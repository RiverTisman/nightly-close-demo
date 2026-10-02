/* Demo-only: plain-language explainer bubbles next to each feature, plus
   the step-by-step guide. Nothing here changes what the app computes.

   Bubbles are attached by matching visible text on the page, so the real
   templates stay untouched. Viewers can hide them with the switch in the
   demo banner; the choice is remembered in this browser only. */
(function () {
  var TIPS = [
    // ---- Daily Close form
    { sel: '.demo-load', where: 'after', text: 'Fills in a made-up night for you, so you can try the app without real files.' },
    { sel: '.shift-select', where: 'after', text: 'Pick lunch, dinner, or the whole day. The app only shows people who worked that part of the day.' },
    { sel: 'label.date-field .date-field-label', where: 'after', text: 'The night you are closing out. It is used to match time-clock punches to the right day.' },
    { sel: '.ticket-title', match: /^Toast Shift Report$/, where: 'header', text: 'Toast is the register. This file says how much each server made in cash and card tips tonight.' },
    { sel: '.ticket-title', match: /^HotSchedules Roster$/, where: 'header', page: 'daily_close', text: 'HotSchedules is the staff schedule. Toast doesn’t know who worked as a runner, busser or barback, so the app gets that from here.' },
    { sel: '.date-field-label', match: /cut from the tip pool/, where: 'after', text: 'Sent someone home early so they don’t share tips tonight? Type a few letters of their name and pick them from the list.' },
    { sel: '.date-field-label', match: /different support-staff role/, where: 'after', text: 'Someone scheduled as a busser who worked as a runner? Pick them and the role they actually worked, so they’re paid in the right group.' },
    { sel: '.fire-button', where: 'after', text: 'Runs every check and builds the lists you paste into the tip sheet.' },

    // ---- Punch Report Check form
    { sel: '.ticket-title', match: /^ADP Punch Report$/, where: 'header', text: 'ADP is the time clock. This file shows when everyone clocked in and out.' },
    { sel: '.ticket-title', match: /^Tip Sheet \(previous day\)$/, where: 'header', text: 'The finished tip sheet from that day. The app checks it against the time clock to make sure everyone is paid correctly.' },
    { sel: '.ticket-title', match: /^HotSchedules Roster$/, where: 'header', page: 'punch_check', text: 'Optional. If someone forgot to clock out, this lets the app suggest a time based on when their coworkers left.' },

    // ---- Daily Close results
    { sel: '.ticket-title', match: /^Cut confirmation$/, where: 'header', text: 'Compares the people you said were cut with what the data shows, so nobody is cut or missed by accident.' },
    { sel: '.ticket-title', match: /^Server tip entries$/, where: 'header', text: 'Each server’s cash and card tips, already written the way the tip sheet needs the names (LAST, FIRST).' },
    { sel: '.server-copy-btn', where: 'buttons', text: 'Copies these rows so you can paste them all at once into the tip sheet. Anything flagged under the 4-hour minimum is left out until you check it.' },
    { sel: '.group-label', match: /^Master Server Cash Sales Sheet/, where: 'after', text: 'Copies server numbers and cash amounts for the cash sales sheet. Numbers and cash are separate copies because the name column in between fills itself in.' },
    { sel: '.flag-pill', where: 'table', once: true, text: 'Rows marked REVIEW need a person to decide. The app never guesses: the reason is written right under the row.' },
    { sel: '.gratuity-resolve', where: 'before', once: true, text: 'A big party’s automatic tip goes to the server by default. If the house rule is to share it with the bar, split it here.' },
    { sel: '.under4h-resolve', where: 'before', once: true, text: 'Dinner shifts under 4 hours don’t share in the tip pool. If this person should be included anyway, add them back.' },
    { sel: '.section-label', match: /^Bartender pool totals/, where: 'after', text: 'Bartenders all use one shared register, so Toast only knows the total. You type in which bartenders worked.' },
    { sel: '.group-label', match: /^Anomalies flagged/, where: 'after', text: 'Data that doesn’t add up, like hours on the clock with no sales. It usually means someone clocked out wrong.' },
    { sel: '.group-label', match: /^Unmatched Toast names/, where: 'after', text: 'A name the app has never seen. It won’t guess who it is. A manager adds it once and the app remembers it.' },
    { sel: '.section-label', match: /^Support staff/, where: 'after', text: 'Runners, bussers, baristas and barbacks aren’t in Toast, so these lists come from the schedule. Copy each one into its own section of the tip sheet.' },
    { sel: '.support-copy-btn', where: 'buttons', text: 'Copies just the names for this role. Paste them into this role’s section of the tip sheet; their share is worked out there.' },
    { sel: '.oncall-resolve', where: 'before', once: true, text: 'Call-ins and trainees stay off the list until you confirm they actually worked. Terminated staff are left off automatically.' },
    { sel: '.ticket-title', match: /^HotSchedules daily roster/, where: 'header', text: 'The full schedule for the day. Tap to open it. Anything unusual is marked REVIEW.' },

    // ---- Punch Report Check results
    { sel: '.ticket-title', match: /^Mapping check against/, where: 'header', text: 'Checks that every name is spelled exactly like the tip sheet’s list. One extra space and that person’s tips quietly add up to $0.' },
    { sel: '.ticket-title', match: /^ADP clock-in\/out check/, where: 'header', text: 'Everyone who clocked in that day. Missing clock-outs and short dinner shifts are flagged.' },
    { sel: '.ticket-title', match: /^Recommended clock-out times/, where: 'header', text: 'For someone who forgot to clock out: a suggested time, based on when coworkers in the same role and shift left. Check with the person first.' },
    { sel: '.ticket-title', match: /^Possible early cuts/, where: 'header', text: 'People who left 2 or more hours before their coworkers. They may need a smaller share of tips.' },
    { sel: '.ticket-title', match: /^Tip Sheet vs\. ADP/, where: 'header', text: 'On the tip sheet but never clocked in, clocked in but missing from the sheet, or listed in the wrong group. Each one means someone gets paid wrong.' }
  ];

  var KEY = 'nc_demo_explain_off';
  function isOff() { try { return localStorage.getItem(KEY) === '1'; } catch (e) { return false; } }
  function setOff(v) { try { localStorage.setItem(KEY, v ? '1' : '0'); } catch (e) {} }

  function bubble(text) {
    var d = document.createElement('div');
    d.className = 'explain-bubble';
    d.setAttribute('role', 'note');
    d.innerHTML = '<span class="explain-icon" aria-hidden="true"></span><span class="explain-text"></span>';
    d.querySelector('.explain-text').textContent = text;
    return d;
  }

  function place(el, where, b) {
    if (where === 'header') {
      var h = el.closest('.ticket-header');
      if (h && h.tagName === 'SUMMARY') { h.parentNode.insertBefore(b, h.nextSibling); return; }
      el = h || el;
      el.parentNode.insertBefore(b, el.nextSibling);
    } else if (where === 'buttons') {
      var g = el.closest('.copy-buttons') || el;
      g.parentNode.insertBefore(b, g.nextSibling);
    } else if (where === 'table') {
      var tbl = el.closest('table') || el;
      tbl.parentNode.insertBefore(b, tbl);
    } else if (where === 'before') {
      el.parentNode.insertBefore(b, el);
    } else {
      if (el.classList.contains('date-field-label')) b.classList.add('explain-flush');
      el.parentNode.insertBefore(b, el.nextSibling);
    }
  }

  function apply() {
    var page = document.body.getAttribute('data-page') || '';
    TIPS.forEach(function (t) {
      if (t.page && t.page !== page) return;
      var els = Array.prototype.slice.call(document.querySelectorAll(t.sel)).filter(function (el) {
        return !t.match || t.match.test(el.textContent.trim());
      });
      if (t.once || t.where === 'buttons') els = els.slice(0, 1);
      els.forEach(function (el) { place(el, t.where, bubble(t.text)); });
    });
  }

  function syncToggle() {
    var off = isOff();
    document.documentElement.classList.toggle('explain-hidden', off);
    var btn = document.getElementById('explain-toggle');
    if (btn) {
      btn.setAttribute('aria-pressed', off ? 'false' : 'true');
      btn.textContent = off ? 'Show explanations' : 'Hide explanations';
    }
  }

  document.addEventListener('DOMContentLoaded', function () {
    apply();
    var btn = document.getElementById('explain-toggle');
    if (btn) btn.addEventListener('click', function () { setOff(!isOff()); syncToggle(); });
    syncToggle();
  });
})();
