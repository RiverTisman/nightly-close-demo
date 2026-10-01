/* Remembers the last date entered into a "roster_date" field for 3 hours,
   so running several reports back to back (AM close, then PM close off a
   different upload, then a Punch Report Check on the same night) doesn't
   require retyping the date every time. Shared across all three forms via
   one localStorage key -- it's the same real-world "date being closed,"
   not a per-page concept.

   The 3-hour window slides forward on every use (see markUsed below),
   not just from the first entry: a shift that spans AM and PM closes a
   few hours apart should keep working without the date going stale
   halfway through. Stop touching the app for 3 hours and it expires, so
   a new day never silently inherits yesterday's date. */
(function () {
  var KEY = 'avra_last_roster_date';
  var TTL_MS = 3 * 60 * 60 * 1000;

  function read() {
    try {
      var raw = localStorage.getItem(KEY);
      if (!raw) return null;
      var data = JSON.parse(raw);
      if (!data || !data.value || (Date.now() - data.savedAt) >= TTL_MS) return null;
      return data.value;
    } catch (e) {
      return null;
    }
  }

  function write(value) {
    try {
      localStorage.setItem(KEY, JSON.stringify({ value: value, savedAt: Date.now() }));
    } catch (e) {
      // localStorage unavailable (private browsing, quota) -- fine, this
      // is a convenience, not something the form depends on.
    }
  }

  function init() {
    document.querySelectorAll('input[type="date"][name="roster_date"]').forEach(function (input) {
      if (!input.value) {
        var remembered = read();
        if (remembered) input.value = remembered;
      }

      input.addEventListener('change', function () {
        if (input.value) write(input.value);
      });

      var form = input.closest('form');
      if (form) {
        form.addEventListener('submit', function () {
          if (input.value) write(input.value);
        });
      }
    });
  }

  window.AvraDateMemory = { init: init };
})();
