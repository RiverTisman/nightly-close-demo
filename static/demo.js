/* Demo-only: one click fills a form with the fictional sample exports
   from /sample, exactly as if a manager had picked them by hand -- the
   files go into the real <input type="file"> elements via DataTransfer,
   so the normal form submit and the real server-side parsing run
   unchanged. */
(function () {
  var DEMO_DATE = '2026-09-24';

  function loadInto(input, url, filename) {
    return fetch(url).then(function (r) {
      if (!r.ok) throw new Error('Could not load ' + filename);
      return r.blob();
    }).then(function (blob) {
      var dt = new DataTransfer();
      dt.items.add(new File([blob], filename, { type: blob.type }));
      input.files = dt.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }

  window.NightlyCloseDemo = {
    init: function (buttonId, files) {
      var btn = document.getElementById(buttonId);
      if (!btn) return;
      btn.addEventListener('click', function () {
        var form = btn.closest('form') || document.querySelector('form.ticket');
        btn.disabled = true;
        btn.textContent = 'Loading sample files…';
        var jobs = Object.keys(files).map(function (field) {
          var input = form.querySelector('input[name="' + field + '"]');
          return input ? loadInto(input, '/sample/' + files[field], files[field]) : Promise.resolve();
        });
        var date = form.querySelector('input[name="roster_date"]');
        if (date) {
          date.value = DEMO_DATE;
          date.dispatchEvent(new Event('change', { bubbles: true }));
        }
        Promise.all(jobs).then(function () {
          btn.textContent = 'Sample files loaded — press Fire';
          btn.classList.add('demo-loaded');
          var fire = form.querySelector('.fire-button');
          if (fire) fire.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }).catch(function (e) {
          btn.disabled = false;
          btn.textContent = 'Load sample files';
          alert(e.message);
        });
      });
    }
  };
})();
