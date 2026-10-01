/* Small presentational-only helpers -- no business logic, nothing here
   changes what gets submitted or computed. */
(function () {
  function initShiftPills() {
    var groups = {};
    document.querySelectorAll('.shift-select input[type="radio"]').forEach(function (radio) {
      groups[radio.name] = groups[radio.name] || [];
      groups[radio.name].push(radio);
    });
    Object.keys(groups).forEach(function (name) {
      var radios = groups[name];
      function sync() {
        radios.forEach(function (r) {
          r.closest('label').classList.toggle('is-active', r.checked);
        });
      }
      radios.forEach(function (r) { r.addEventListener('change', sync); });
      sync();
    });
  }
  document.addEventListener('DOMContentLoaded', initShiftPills);
})();
