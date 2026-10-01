/* Reusable search-as-you-type staff picker: type a few letters of a name,
   pick from the matching real employees, get a removable chip. Exists to
   kill a whole bug class -- free-typed names into a plain textarea (the
   old "cut list"/"role swap" boxes) never reliably matched a real
   employee (typos, nicknames, "Last, First" vs "First Last", middle
   initials...), no matter how much the server-side fuzzy-matching got
   patched. Selecting from a real list instead of typing free text makes a
   mismatch structurally impossible: what gets submitted is always exactly
   what was offered, never text to re-interpret.

   Deliberately generic -- nothing here is specific to "who got cut" or
   "role swaps". A mount just needs a staff list (any array of objects
   with a label field) and a container; a different page can hand it a
   different, smaller staff list (e.g. servers only) without touching this
   file. Usage:

     AvraStaffPicker.init(document.getElementById('cut-picker'), staffList, {
       hiddenInputName: 'cut_employees_json',
       placeholder: 'Type a name...',
     });

     AvraStaffPicker.init(document.getElementById('role-swap-picker'), staffList, {
       hiddenInputName: 'role_swap_json',
       placeholder: 'Type a name...',
       extraField: {
         key: 'role',
         options: [{value: 'Busser', label: 'Busser'}, ...],
       },
     });

   Every mount's current selection is always available as JSON in its
   hidden <input>, so it submits with the surrounding <form> like any
   other field -- no extra wiring needed at submit time.

   opts.maxSelected caps the chip count -- picking a new person past the
   cap evicts the oldest chip first (so maxSelected: 1 behaves like a
   single-select that always holds "whoever you picked most recently").
   opts.onChange(selected), if given, fires after every add/remove/
   extra-field edit with a shallow copy of the current selection --
   useful when another part of the page needs to react to who's picked
   (e.g. populating an edit form with the selected person's details). */
(function () {
  function normalize(s) {
    return (s || '').toString().toLowerCase().trim();
  }

  function init(container, staffList, opts) {
    opts = opts || {};
    var labelKey = opts.labelKey || 'canonical';
    var valueKey = opts.valueKey || 'canonical';
    var maxSuggestions = opts.maxSuggestions || 8;
    var extraField = opts.extraField || null;
    var selected = (opts.initial || []).slice();

    container.classList.add('staff-picker');
    container.innerHTML = '';

    var inputWrap = document.createElement('div');
    inputWrap.className = 'staff-picker-input-wrap';

    var input = document.createElement('input');
    input.type = 'text';
    input.className = 'staff-picker-input';
    input.setAttribute('autocomplete', 'off');
    input.placeholder = opts.placeholder || 'Type a name...';

    var dropdown = document.createElement('div');
    dropdown.className = 'staff-picker-dropdown';
    dropdown.hidden = true;

    inputWrap.appendChild(input);
    inputWrap.appendChild(dropdown);

    var chipsEl = document.createElement('div');
    chipsEl.className = 'staff-picker-chips';

    var hiddenInput = document.createElement('input');
    hiddenInput.type = 'hidden';
    hiddenInput.name = opts.hiddenInputName;

    container.appendChild(inputWrap);
    container.appendChild(chipsEl);
    container.appendChild(hiddenInput);

    var currentMatches = [];
    var highlightedIdx = -1;

    function serialize() {
      hiddenInput.value = JSON.stringify(selected.map(function (item) {
        var out = {};
        out[valueKey] = item[valueKey];
        if (extraField) out[extraField.key] = item[extraField.key];
        return out;
      }));
      if (typeof opts.onChange === 'function') opts.onChange(selected.slice());
    }

    function isSelected(item) {
      return selected.some(function (s) { return s[valueKey] === item[valueKey]; });
    }

    function renderChips() {
      chipsEl.innerHTML = '';
      selected.forEach(function (item, idx) {
        var chip = document.createElement('span');
        chip.className = 'staff-picker-chip';

        var nameEl = document.createElement('span');
        nameEl.className = 'staff-picker-chip-name';
        nameEl.textContent = item[labelKey];
        chip.appendChild(nameEl);

        if (extraField) {
          var select = document.createElement('select');
          select.className = 'staff-picker-chip-role';
          extraField.options.forEach(function (opt) {
            var o = document.createElement('option');
            o.value = opt.value;
            o.textContent = opt.label;
            if (item[extraField.key] === opt.value) o.selected = true;
            select.appendChild(o);
          });
          if (item[extraField.key] === undefined) {
            item[extraField.key] = extraField.options[0].value;
            select.value = item[extraField.key];
          }
          select.addEventListener('change', function () {
            item[extraField.key] = select.value;
            serialize();
          });
          chip.appendChild(select);
        }

        var removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'staff-picker-chip-remove';
        removeBtn.setAttribute('aria-label', 'Remove ' + item[labelKey]);
        removeBtn.textContent = '×';
        removeBtn.addEventListener('click', function () {
          selected.splice(idx, 1);
          renderChips();
          serialize();
        });
        chip.appendChild(removeBtn);

        chipsEl.appendChild(chip);
      });
    }

    function closeDropdown() {
      dropdown.hidden = true;
      dropdown.innerHTML = '';
      currentMatches = [];
      highlightedIdx = -1;
    }

    function selectItem(item) {
      if (opts.maxSelected && selected.length >= opts.maxSelected) {
        selected.splice(0, selected.length - opts.maxSelected + 1);
      }
      selected.push(Object.assign({}, item));
      renderChips();
      serialize();
      input.value = '';
      closeDropdown();
      input.focus();
    }

    function renderDropdown(query) {
      var matches = staffList
        .filter(function (item) { return !isSelected(item); })
        .filter(function (item) { return normalize(item[labelKey]).indexOf(query) !== -1; })
        .slice(0, maxSuggestions);
      currentMatches = matches;
      highlightedIdx = -1;
      dropdown.innerHTML = '';

      if (matches.length === 0) {
        var empty = document.createElement('div');
        empty.className = 'staff-picker-empty';
        empty.textContent = 'No matches';
        dropdown.appendChild(empty);
        dropdown.hidden = false;
        return;
      }

      matches.forEach(function (item, idx) {
        var opt = document.createElement('div');
        opt.className = 'staff-picker-option';
        opt.textContent = item[labelKey];
        opt.dataset.idx = idx;
        opt.addEventListener('mousedown', function (e) {
          // mousedown (not click) fires before the input's blur, so the
          // dropdown doesn't close itself out from under the click.
          e.preventDefault();
          selectItem(item);
        });
        opt.addEventListener('mouseenter', function () {
          setHighlighted(idx);
        });
        dropdown.appendChild(opt);
      });
      dropdown.hidden = false;
    }

    function setHighlighted(idx) {
      var opts = dropdown.querySelectorAll('.staff-picker-option');
      opts.forEach(function (el) { el.classList.remove('highlighted'); });
      if (idx >= 0 && idx < opts.length) {
        opts[idx].classList.add('highlighted');
        highlightedIdx = idx;
      } else {
        highlightedIdx = -1;
      }
    }

    input.addEventListener('input', function () {
      var query = normalize(input.value);
      if (!query) {
        closeDropdown();
        return;
      }
      renderDropdown(query);
    });

    input.addEventListener('keydown', function (e) {
      if (dropdown.hidden) return;
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setHighlighted(Math.min(highlightedIdx + 1, currentMatches.length - 1));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setHighlighted(Math.max(highlightedIdx - 1, 0));
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (highlightedIdx >= 0 && currentMatches[highlightedIdx]) {
          selectItem(currentMatches[highlightedIdx]);
        } else if (currentMatches.length === 1) {
          selectItem(currentMatches[0]);
        }
      } else if (e.key === 'Escape') {
        closeDropdown();
      } else if (e.key === 'Backspace' && !input.value && selected.length) {
        selected.pop();
        renderChips();
        serialize();
      }
    });

    input.addEventListener('blur', function () {
      // Delay so a mousedown-driven selectItem() (above) still fires first.
      setTimeout(closeDropdown, 150);
    });

    renderChips();
    serialize();
  }

  window.AvraStaffPicker = { init: init };
})();
