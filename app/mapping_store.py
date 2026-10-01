"""
Makes employee_mapping.json survive a restart.

The bug this fixes
------------------
Every page in this app reads employee_mapping.json off local disk, and
/add-employee wrote it back to local disk. That works perfectly on a
laptop and not at all on Render: the container's filesystem is rebuilt
from the image on every deploy AND on every wake from the free tier's
15-minute idle spin-down. So an employee added through the website
existed for as long as the container did -- often under an hour -- and
then silently reverted to whatever was committed in git.

That is why "adding an employee hasn't worked." A new hire was added
through the form, the form said "Added", and he was genuinely gone by
the next night's close, because the only durable copy of the mapping is
the one in the GitHub repo.

So the GitHub repo is now treated as the real store:

  * On startup, pull employee_mapping.json from the repo over the
    container's copy, so a fresh container starts from the newest
    mapping rather than from whenever the image was built.
  * On every write, commit the new file back to the repo. That both
    persists it and gives a full audit trail of who was added when --
    something the local-file version never had.
  * Pushing to `main` also triggers Render's auto-deploy, so the running
    app picks up its own change. Restarting mid-write is safe: the
    commit lands first and the local file is already correct.

If no token is configured the app still works exactly as before -- local
file only -- but says so plainly on the page instead of implying the
change is permanent. Silence about a change that's about to vanish is
the failure mode this module exists to remove.

Configuration (Render dashboard -> Environment):

  AVRA_GITHUB_TOKEN   a fine-grained personal access token with
                      Contents: Read and write on your-org/your-repo
  AVRA_GITHUB_REPO    owner/repo      (default your-org/your-repo)
  AVRA_GITHUB_BRANCH  branch to commit to (default main)
  AVRA_MAPPING_PATH   where the local copy lives (default: repo root)
"""

import base64
import json
import os
import urllib.error
import urllib.request
from pathlib import Path

GITHUB_API = "https://api.github.com"
DEFAULT_REPO = "your-org/your-repo"
DEFAULT_BRANCH = "main"
MAPPING_FILENAME = "employee_mapping.json"

# Set by the most recent pull/push so the UI can report what actually
# happened rather than guessing. Not state the app depends on -- purely
# for display.
_last_sync = {"ok": None, "detail": "No GitHub sync attempted yet this session."}


TOKEN_KEY = "AVRA_GITHUB_TOKEN"

# Where Render mounts Secret Files. Overridable mainly so this is
# testable without pretending to be a container.
SECRETS_DIR = os.environ.get("AVRA_SECRETS_DIR") or "/etc/secrets"

# GitHub's own token prefixes: fine-grained PAT, classic PAT, OAuth,
# user-to-server, server-to-server, refresh. Used only to recognise
# which file in the secrets directory is the token -- never to validate
# one, which is GitHub's job.
_TOKEN_PREFIXES = ("github_pat_", "ghp_", "gho_", "ghu_", "ghs_", "ghr_")


def _token_from_env():
    """The token from an environment variable, tolerating a key name that
    picked up stray whitespace on its way into the dashboard.

    Pasting a variable name into a web form very easily carries a
    trailing space or a non-breaking space with it, and the result is
    invisible: the dashboard renders "AVRA_GITHUB_TOKEN " exactly like
    "AVRA_GITHUB_TOKEN", so the setting looks correct while
    os.environ.get() misses it entirely. Matching on the stripped key
    name costs nothing and can't select a different variable.
    """
    direct = os.environ.get(TOKEN_KEY)
    if direct and direct.strip():
        return direct.strip(), TOKEN_KEY
    for key, value in os.environ.items():
        if key.strip() == TOKEN_KEY and value and value.strip():
            return value.strip(), key
    return "", None


def _token_from_secret_file():
    """The token from Render's Secret Files, if that's where it was put.

    "Environment Variables" and "Secret Files" sit next to each other in
    Render's Environment tab, and choosing the second one is an entirely
    reasonable reading of "store a secret". But a Secret File is mounted
    at /etc/secrets/<name> and never becomes an environment variable, so
    the setting looks perfectly correct on screen while the app sees
    nothing at all -- with no error anywhere to suggest why.

    Rather than make someone deduce that, the token is simply read from
    there too: first by the expected filename, then by looking for a
    file whose contents carry a GitHub token prefix. This reads only the
    service's own secrets directory, and matches on the value's shape so
    an unrelated secret can't be mistaken for the token and sent to
    GitHub.
    """
    directory = Path(SECRETS_DIR)
    if not directory.is_dir():
        return "", None

    named = directory / TOKEN_KEY
    if named.is_file():
        try:
            value = named.read_text(encoding="utf-8").strip()
        except OSError:
            value = ""
        if value:
            return value, str(named)

    try:
        candidates = sorted(p for p in directory.iterdir() if p.is_file())
    except OSError:
        return "", None
    for path in candidates:
        try:
            value = path.read_text(encoding="utf-8").strip()
        except (OSError, UnicodeDecodeError):
            continue
        if value.startswith(_TOKEN_PREFIXES):
            return value, str(path)
    return "", None


def _token_with_source():
    value, source = _token_from_env()
    if value:
        return value, source
    return _token_from_secret_file()


def _token():
    return _token_with_source()[0]


def _repo():
    return (os.environ.get("AVRA_GITHUB_REPO") or DEFAULT_REPO).strip()


def _branch():
    return (os.environ.get("AVRA_GITHUB_BRANCH") or DEFAULT_BRANCH).strip()


def is_configured():
    return bool(_token())


def _contents_url():
    return f"{GITHUB_API}/repos/{_repo()}/contents/{MAPPING_FILENAME}"


def _request(url, method="GET", payload=None):
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=body, method=method)
    req.add_header("Authorization", f"Bearer {_token()}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    req.add_header("User-Agent", "avra-nightly-check")
    if body is not None:
        req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _describe_http_error(e):
    """GitHub's own error message, which is far more actionable than the
    status code alone ("Resource not accessible by personal access
    token" tells the manager exactly what to fix; "403" doesn't)."""
    try:
        detail = json.loads(e.read().decode("utf-8")).get("message", "")
    except Exception:
        detail = ""
    if e.code == 401:
        hint = "the token is wrong or expired"
    elif e.code == 403:
        hint = f"the token can't write to {_repo()} -- it needs Contents: Read and write"
    elif e.code == 404:
        hint = f"{_repo()} or {MAPPING_FILENAME} wasn't found (a token without access reads as 404)"
    elif e.code == 409:
        hint = "the file changed on GitHub since this app last read it"
    else:
        hint = "GitHub rejected the request"
    return f"HTTP {e.code}: {hint}" + (f" -- {detail}" if detail else "")


def _why_no_token():
    """A specific, actionable reason the token isn't visible -- not just
    "it isn't set".

    "Set AVRA_GITHUB_TOKEN in the dashboard" is useless advice to
    somebody who has already done exactly that and is looking at the
    variable on their screen. Several genuinely different faults all
    present as an absent token, and only the process itself can tell
    them apart, so it says which one it is instead of making a person
    guess and re-check.
    """
    env = env_diagnostics()

    if env["exact_match_present"] and env["token_length"] == 0:
        return ("The variable AVRA_GITHUB_TOKEN exists here, but its value is empty. "
                "Re-paste the token into the Value field in Render and save.")

    if env["secret_files_visible"]:
        return (
            f"There are Secret Files on this service ({', '.join(env['secret_files_visible'])}), "
            "but none of them contains something shaped like a GitHub token (they all start "
            f"{'/'.join(_TOKEN_PREFIXES[:2])}...). If one of those is meant to be the token, "
            "check its contents were saved. Otherwise add the token under Environment "
            f"Variables as {TOKEN_KEY}."
        )

    visible = [k for k in env["similar_keys_visible"] if k.strip() != TOKEN_KEY]
    if not env["similar_keys_visible"]:
        return (
            f"This service ({os.environ.get('RENDER_SERVICE_NAME') or 'this app'}) is receiving no "
            "AVRA_* environment variable and has no Secret Files at all, so the setting isn't "
            "reaching this service from anywhere. The app reads both, so this isn't the "
            "Environment-Variables-vs-Secret-Files distinction. What's left: it's saved on a "
            "DIFFERENT Render service, or it's in an Environment Group that was never linked to "
            "this service (Environment tab -> \"Linked Environment Groups\"), or it was typed "
            "but never saved. Confirm the service name above is the one you edited."
        )
    return (
        f"This app can see {', '.join(visible)} from Render, but not AVRA_GITHUB_TOKEN. "
        "So the environment is reaching the service and this one variable specifically is "
        "missing or misnamed -- check it's on the same service as the others and spelled "
        "AVRA_GITHUB_TOKEN exactly."
    )


def deployment_info():
    """Which build is actually running, from the variables Render sets
    itself. Answers "did my change deploy?" without a dashboard."""
    return {
        "service": os.environ.get("RENDER_SERVICE_NAME"),
        "commit": (os.environ.get("RENDER_GIT_COMMIT") or "")[:7] or None,
        "branch": os.environ.get("RENDER_GIT_BRANCH"),
        "url": os.environ.get("RENDER_EXTERNAL_URL"),
        "on_render": bool(os.environ.get("RENDER")),
    }


def status():
    """What the page should tell the manager about durability."""
    if not is_configured():
        return {
            "configured": False,
            "ok": False,
            "detail": (
                "No GitHub token configured, so changes are saved only to this server's own disk. "
                "On Render's free tier that disk is wiped every time the service redeploys or "
                "wakes from its 15-minute idle sleep -- a change made here can disappear within "
                "the hour."
            ),
            "diagnosis": _why_no_token(),
            "repo": _repo(),
            "branch": _branch(),
            "deployment": deployment_info(),
        }
    env = env_diagnostics()
    note = ""
    if env["matched_via_padded_key"]:
        note = (f" (Found it under the key {env['matched_via_padded_key'][0]!r} -- that name has "
                "stray whitespace in it, which is why it looked correct in Render but wasn't "
                "being picked up. It's being used anyway; retype the Key without the space when "
                "convenient.)")
    elif env["token_source"] and env["token_source"].startswith(SECRETS_DIR):
        note = (f" (Read from the Secret File {env['token_source']!r} rather than an environment "
                "variable. That works and needs no change.)")
    return {"configured": True, "ok": _last_sync["ok"],
            "detail": _last_sync["detail"] + note,
            "repo": _repo(), "branch": _branch(), "deployment": deployment_info(),
            "diagnosis": None}


def _remember(ok, detail):
    _last_sync["ok"] = ok
    _last_sync["detail"] = detail
    return {"ok": ok, "detail": detail}


def verify_write_access():
    """Confirms the token can actually WRITE, not merely read.

    Reading and writing are separate permissions, and only one of them
    is exercised at startup -- so a token granted Contents: Read (an
    easy mistake, it's the default) pulls the mapping perfectly, shows
    every green light, and then fails the first time someone saves. That
    is precisely the silent-failure shape this module exists to remove,
    so the capability is checked directly instead of being inferred from
    a successful read.

    GET /repos/{owner}/{repo} returns a `permissions` block for an
    authenticated caller; `permissions.push` is the flag that governs
    committing a file. One cheap call at startup, and the answer is
    definitive rather than hopeful.
    """
    if not is_configured():
        return None, "No token configured."
    try:
        repo = _request(f"{GITHUB_API}/repos/{_repo()}")
    except urllib.error.HTTPError as e:
        return False, _describe_http_error(e)
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

    perms = repo.get("permissions") or {}
    if perms.get("push"):
        return True, f"Token can write to {_repo()}."
    return False, (
        f"The token can read {_repo()} but NOT write to it. Saves will fail. In the token's "
        "settings, set Repository permissions -> Contents to \"Read and write\" (not just "
        "\"Read\"), then restart the Render service."
    )


def env_diagnostics():
    """What the running process can actually see, for when the token is
    "definitely set" and the app disagrees.

    Reports env var NAMES and value LENGTHS only -- never any part of a
    value, since this is served on an unauthenticated endpoint. That is
    enough to separate the four things that all present identically as
    "no token configured":

      * nothing listed          -> Render isn't passing the variable to
                                   this process at all: wrong service,
                                   saved as a Secret File (which mounts
                                   a file rather than setting an env
                                   var), or saved without a restart.
      * a name that looks right -> copy-paste picked up a trailing space
        but doesn't match          or non-breaking space in the KEY.
                                   `exact_match` is the authoritative
                                   answer, not how the name looks.
      * length 0                -> the variable exists but is empty.
      * exact_match true, and   -> the fault is in this app, not the
        the app still complains    configuration.
    """
    interesting = {k: len(v) for k, v in os.environ.items()
                   if "AVRA" in k.upper() or "GITHUB" in k.upper()}
    padded = [k for k in os.environ if k != TOKEN_KEY and k.strip() == TOKEN_KEY]
    try:
        secret_files = sorted(p.name for p in Path(SECRETS_DIR).iterdir() if p.is_file())
    except OSError:
        secret_files = []
    value, source = _token_with_source()
    return {
        "expected_key": TOKEN_KEY,
        "exact_match_present": TOKEN_KEY in os.environ,
        "matched_via_padded_key": padded,
        "token_source": source,
        "secrets_dir": SECRETS_DIR,
        "secrets_dir_exists": Path(SECRETS_DIR).is_dir(),
        "secret_files_visible": secret_files,
        "token_length": len(value),
        "similar_keys_visible": sorted(interesting),
        "similar_key_lengths": {k: interesting[k] for k in sorted(interesting)},
        "total_env_vars": len(os.environ),
    }


def pull_to_local(local_path):
    """Overwrites the local mapping file with the repo's copy. Called once
    at startup.

    A failure here is deliberately NOT fatal: the container still has the
    copy baked into its image, which is stale at worst, and refusing to
    start the whole app over a GitHub hiccup would take down the nightly
    close for a problem that only affects newly added employees. The
    failure is recorded and shown on the Add Employee page instead.
    """
    if not is_configured():
        return _remember(None, "GitHub sync is off -- using this server's local file only.")
    try:
        data = _request(_contents_url() + f"?ref={_branch()}")
        content = base64.b64decode(data["content"]).decode("utf-8")
        json.loads(content)  # never overwrite a working local file with something unparseable
        Path(local_path).write_text(content, encoding="utf-8")
    except urllib.error.HTTPError as e:
        return _remember(False, f"Couldn't load the mapping from GitHub -- {_describe_http_error(e)}. "
                                "Using the copy built into this deploy, which may be out of date.")
    except Exception as e:
        return _remember(False, f"Couldn't load the mapping from GitHub ({type(e).__name__}: {e}). "
                                "Using the copy built into this deploy, which may be out of date.")

    # The read worked; now prove the token can also commit, before the
    # UI implies to anyone that their changes will stick.
    can_write, detail = verify_write_access()
    if can_write:
        return _remember(True, f"Loaded the mapping from {_repo()}@{_branch()} at startup. {detail}")
    return _remember(False, f"Loaded the mapping from {_repo()}@{_branch()}, but {detail}")


def push_text(text, message):
    """Commits `text` as the repo's employee_mapping.json.

    The blob SHA is re-read immediately before the commit rather than
    cached, so two people saving from different browsers produce two
    commits instead of one silently clobbering the other -- and if it
    genuinely races, GitHub returns 409 and the manager is told to redo
    the change rather than being told it saved when it didn't.
    """
    if not is_configured():
        return _remember(None, "Saved to this server's disk only -- no GitHub token is configured, "
                               "so this change will be lost when the service restarts.")
    try:
        sha = None
        try:
            sha = _request(_contents_url() + f"?ref={_branch()}").get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:  # 404 = file not in the repo yet, which a create handles
                raise
        payload = {
            "message": message,
            "content": base64.b64encode(text.encode("utf-8")).decode("ascii"),
            "branch": _branch(),
        }
        if sha:
            payload["sha"] = sha
        result = _request(_contents_url(), method="PUT", payload=payload)
        short = (result.get("commit", {}).get("sha") or "")[:7]
        return _remember(True, f"Committed to {_repo()}@{_branch()}"
                               + (f" as {short}." if short else "."))
    except urllib.error.HTTPError as e:
        return _remember(False, f"Saved on this server, but NOT committed to GitHub -- "
                                f"{_describe_http_error(e)}. This change will be lost when the "
                                "service restarts. Fix the token and make the change again.")
    except Exception as e:
        return _remember(False, f"Saved on this server, but NOT committed to GitHub "
                                f"({type(e).__name__}: {e}). This change will be lost when the "
                                "service restarts.")
