/* Keeps a results page on screen when you switch nav tabs and come back,
   instead of losing it back to the blank upload form -- this app is
   stateless server-side (no DB, no session), so "restoring" a page means
   literally replaying the last rendered results HTML back into the
   browser. Uses sessionStorage (not localStorage) so it only lasts for
   this browser tab/session, not indefinitely across days.

   document.write (not innerHTML) is used to restore because innerHTML
   assignment does not execute embedded <script> tags -- the results
   pages' resolve widgets, copy buttons, etc. all depend on their bottom
   <script> block actually running, not just the markup being present. */
(function () {
  function key(pageId) { return 'avra_results:' + pageId; }

  function save(pageId) {
    try {
      sessionStorage.setItem(key(pageId), document.documentElement.outerHTML);
    } catch (e) {
      // storage full/unavailable -- fine, this is a convenience cache
    }
  }

  function clear(pageId) {
    try { sessionStorage.removeItem(key(pageId)); } catch (e) { /* ignore */ }
  }

  // Returns true if a cached results page was found and restored (in
  // which case the caller should do nothing else -- the whole document
  // just got replaced). Returns false if there was nothing to restore,
  // so the caller should render its normal blank form as usual.
  function restore(pageId) {
    var html;
    try { html = sessionStorage.getItem(key(pageId)); } catch (e) { return false; }
    if (!html) return false;
    document.open();
    document.write(html);
    document.close();
    return true;
  }

  // Call from a results page (not the form page) once rendering is done
  // and there was no error. Saves this page for later restoration, and
  // makes the "run another" link clear the cache first -- otherwise
  // clicking it would just re-restore the same stale results forever,
  // since it navigates back to the form page that tries to restore.
  function initResultsPage(pageId) {
    save(pageId);
    var backLink = document.querySelector('.back-link');
    if (backLink) {
      backLink.addEventListener('click', function () { clear(pageId); });
    }
  }

  window.AvraResultsCache = { save: save, clear: clear, restore: restore, initResultsPage: initResultsPage };
})();
