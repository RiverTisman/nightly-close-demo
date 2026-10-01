/* Remembers file-input selections across page navigations within this
   browser only (nothing is sent anywhere) -- switching between the Daily
   Close / Daily Summary / Punch Report Check tabs no longer loses files
   you'd already picked but hadn't submitted yet. Uses IndexedDB because
   File objects can be stored there directly (sessionStorage/localStorage
   can't hold binary data), and restores them into the real <input
   type="file"> via the DataTransfer API so normal form submission needs
   no other changes. Cleared automatically after a successful submit (see
   AvraFileCache.clear, called from the results pages) so a new night
   never silently reuses yesterday's file. */
(function () {
  var DB_NAME = 'avra_file_cache';
  var STORE = 'files';
  var dbPromise = null;

  function openDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = function () {
        req.result.createObjectStore(STORE, { keyPath: 'key' });
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
    return dbPromise;
  }

  function withStore(mode, fn) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, mode);
        fn(tx.objectStore(STORE));
        tx.oncomplete = function () { resolve(); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function saveFile(key, file) {
    return withStore('readwrite', function (store) {
      store.put({ key: key, name: file.name, file: file, savedAt: Date.now() });
    });
  }

  function loadFile(key) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(STORE, 'readonly');
        var req = tx.objectStore(STORE).get(key);
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function deleteFile(key) {
    return withStore('readwrite', function (store) { store.delete(key); });
  }

  function keyFor(pageId, fieldName) { return pageId + ':' + fieldName; }
  function lastSubmittedKeyFor(pageId, fieldName) { return 'last-submitted:' + pageId + ':' + fieldName; }

  function initPersistentFileInputs(pageId) {
    if (!window.indexedDB) return;
    document.querySelectorAll('input[type="file"][data-persist]').forEach(function (input) {
      var key = keyFor(pageId, input.name);
      var note = document.createElement('p');
      note.className = 'fine-print file-cache-note';
      note.style.cssText = 'margin:-10px 20px 14px;display:none;';
      var dropzone = input.closest('label.dropzone');
      if (!dropzone) return;
      dropzone.insertAdjacentElement('afterend', note);

      // Real drag-and-drop -- the dropzone copy ("Drop the file here...")
      // implies this works, but a native <input type="file"> only opens
      // the picker on click by default; dropping a file onto it does
      // nothing without this. Assigns the dropped file via the same
      // DataTransfer technique used to restore a cached file below, then
      // fires a real "change" event so the existing change listener
      // (saves to cache, updates the visible label) handles it exactly
      // like a normal file-picker selection -- no separate code path.
      dropzone.addEventListener('dragover', function (e) {
        e.preventDefault();
        dropzone.classList.add('drag-active');
      });
      dropzone.addEventListener('dragleave', function () {
        dropzone.classList.remove('drag-active');
      });
      dropzone.addEventListener('drop', function (e) {
        e.preventDefault();
        dropzone.classList.remove('drag-active');
        var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (!file) return;
        try {
          var dt = new DataTransfer();
          dt.items.add(file);
          input.files = dt.files;
          input.dispatchEvent(new Event('change', { bubbles: true }));
        } catch (err) {
          // DataTransfer assignment not supported in this browser -- the
          // native click-to-browse path still works either way.
        }
      });

      // Purely cosmetic: the dropzone's visible label text mirrors
      // whatever the native input already knows (chosen filename or the
      // placeholder), same underlying selection -- no new interaction.
      var faceLabel = dropzone.querySelector('.dropzone-label');
      var placeholderText = faceLabel ? faceLabel.textContent : null;
      function setFace(name) {
        if (faceLabel) faceLabel.textContent = name ? ('Selected: ' + name) : placeholderText;
      }

      function renderNote(name) {
        note.innerHTML = '';
        note.style.display = 'block';
        note.appendChild(document.createTextNode('Using previously selected: ' + name + ' — '));
        var clearLink = document.createElement('a');
        clearLink.href = '#';
        clearLink.textContent = 'clear';
        clearLink.style.color = 'var(--review)';
        clearLink.addEventListener('click', function (e) {
          e.preventDefault();
          deleteFile(key);
          input.value = '';
          note.style.display = 'none';
          setFace(null);
        });
        note.appendChild(clearLink);
      }

      loadFile(key).then(function (record) {
        if (!record || !record.file) return;
        try {
          var dt = new DataTransfer();
          dt.items.add(record.file);
          input.files = dt.files;
          renderNote(record.name);
          setFace(record.name);
        } catch (e) {
          // DataTransfer restore not supported in this browser -- fine,
          // just means this one input won't auto-restore.
        }
      });

      input.addEventListener('change', function () {
        if (input.files && input.files[0]) {
          saveFile(key, input.files[0]);
          note.style.display = 'none';
          setFace(input.files[0].name);
        }
      });

      // A completed run archives its files under a separate "last
      // submitted" key (see clearFileCache below) instead of just
      // deleting them, specifically so someone closing AM then PM off the
      // same night's uploads doesn't have to re-browse for the same
      // files each time -- but reusing them is always a deliberate click
      // here, never automatic, so a new night can never silently inherit
      // yesterday's file the way plain auto-restore would risk.
      var lastKey = lastSubmittedKeyFor(pageId, input.name);
      loadFile(lastKey).then(function (record) {
        if (!record || !record.file) return;
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'use-previous-upload-btn';
        btn.style.cssText = 'margin:-8px 20px 14px;font-size:13px;padding:4px 10px;border-radius:8px;border:1px solid var(--field-border);background:#fff;cursor:pointer;color:var(--body-text);';
        btn.textContent = 'Use previous upload (' + record.name + ')';
        note.insertAdjacentElement('afterend', btn);
        btn.addEventListener('click', function () {
          try {
            var dt = new DataTransfer();
            dt.items.add(record.file);
            input.files = dt.files;
            saveFile(key, record.file);
            renderNote(record.name);
            setFace(record.name);
          } catch (e) {
            // DataTransfer restore not supported in this browser.
          }
        });
      });
    });
  }

  function clearFileCache(pageId, fieldNames) {
    fieldNames.forEach(function (name) {
      var key = keyFor(pageId, name);
      loadFile(key).then(function (record) {
        if (record && record.file) {
          saveFile(lastSubmittedKeyFor(pageId, name), record.file);
        }
        deleteFile(key);
      });
    });
  }

  window.AvraFileCache = { init: initPersistentFileInputs, clear: clearFileCache };
})();
