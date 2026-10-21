/**
 * Trail Camera Reviewer - Google Drive edition (revised)
 *
 * Deploy as a Web App, "Execute as: Me", "Who has access: Anyone within <your domain>".
 *
 * What changed in this revision (see the notes at the bottom for setup):
 *  - Security: no OAuth token or file access before login; every call after login needs a
 *    server-side session; folder/file IDs are never trusted from the browser; PIN attempts
 *    are throttled.
 *  - Speed: folder IDs are resolved once and remembered; species list is cached; tags are
 *    processed in batches with one Drive call (move + rename) per file.
 *  - Consistency: cameras are identified by Unique Cam ID (not sheet row number); claims are
 *    locked; the media folder is matched on camera number AND dates; retries are idempotent;
 *    Master Data rows carry the Drive file ID so Undo finds the right row.
 */

const CFG = {
  APP_TITLE: 'Trail Camera Reviewer',
  MEDIA_EXT: /\.(png|jpe?g|mp4|avi|mov)$/i,
  VIDEO_EXT: /\.(mp4|avi|mov)$/i,
  STAMP_FMT: 'yyyy:MM:dd HH:mm:ss',
  SESSION_TTL: 21600,      // seconds (CacheService maximum is 6 hours)
  MAX_FAILS: 5,            // failed PIN attempts per student ...
  FAIL_WINDOW: 900,        // ... within this many seconds
  MAX_BATCH: 25,           // most tags accepted in a single call
  MASTER_HEADERS: [
    'Original_Path', 'Folder', 'Reviewer', 'Site', 'Unique_Cam_ID', 'Camera_Number',
    'File_Name', 'Species', 'GPS_Coordinates', 'Original_Time', 'Corrected_Time',
    'Time_Status', 'Offset_Seconds', 'File_ID', 'Stars', 'ID_Doubt', 'Suggested_ID'
  ],
  FILE_ID_COL: 14,         // column numbers in Master Data
  STARS_COL: 15,           // number of students who starred the photo
  DOUBT_COL: 16,           // number of students who clicked "Nahhh" (doubt the ID)
  SUGG_COL: 17,            // different suggested IDs, comma separated
  FEEDBACK_SHEET: 'Photo Feedback',   // one row per student per photo (who starred / doubted / suggested)
  FEEDBACK_HEADERS: ['File_ID', 'Unique_Cam_ID', 'Student', 'Star', 'Nahh', 'Suggested_ID', 'Updated'],
  // Optional: instructors by name (case-insensitive; "jane" also matches "Jane Smith").
  // The easier way is to put "yes" in the Instructor column of the Student Names tab (run setupInstructorColumn once).
  ADMIN_NAMES: [],
  MAX_SUGGESTIONS: 5,
  FOLDERS_SHEET: 'Camera Folders',    // folders chosen in the instructor "Camera setup" tab (new tab, created on first save)
  FOLDERS_HEADERS: ['Scope', 'Key', 'Photos_Folder_ID', 'Photos_Folder_Name', 'Tagged_Folder_ID', 'Tagged_Folder_Name', 'Updated', 'Updated_By'],
  LOG_SHEET: 'Setup Change Log',      // every change made from the instructor tabs (new tab, created on first save)
  LOG_HEADERS: ['Time', 'Instructor', 'Action', 'Camera', 'Detail', 'Old_Value', 'New_Value'],
};
const FOLDER_MIME = 'application/vnd.google-apps.folder';
const DEFAULT_SPECIES = ['|||MAMMALS|||', 'Bobcat', 'Coyote', 'Gray Fox', '|||BIRDS|||', 'Wild Turkey', '|||OTHER|||', 'Human', 'Vehicle'];

/* ------------------------------------------------------------------ */
/* Small utilities                                                     */
/* ------------------------------------------------------------------ */

function cache_() { return CacheService.getScriptCache(); }
function props_() { return PropertiesService.getScriptProperties(); }

function hash_(s) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s));
  return Utilities.base64EncodeWebSafe(d).slice(0, 22);
}

function norm_(s) { return String(s).toLowerCase().replace(/\s+/g, ''); }

function escapeRx_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function withRetry_(fn, attempts) {
  let last;
  for (let i = 0; i < attempts; i++) {
    try { return fn(); } catch (e) { last = e; if (i < attempts - 1) Utilities.sleep(600); }
  }
  throw last;
}

let SS_MEMO_ = null;
// Strictly binds to the container spreadsheet to prevent naming collisions
function getSpreadsheet_() {
  if (SS_MEMO_) return SS_MEMO_;
  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSs) { SS_MEMO_ = activeSs; return activeSs; }
  const files = DriveApp.getFilesByName('Camera Review Sign Ups');
  while (files.hasNext()) {
    const file = files.next();
    if (!file.isTrashed()) { SS_MEMO_ = SpreadsheetApp.open(file); return SS_MEMO_; }
  }
  throw new Error('Cannot find Camera Review Sign Ups sheet');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CFG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ------------------------------------------------------------------ */
/* Semester / date helpers                                             */
/* ------------------------------------------------------------------ */

function semesterKeys_(semString) {
  const target = norm_(semString);
  let mapped = target;
  const m = String(semString).trim().match(/^(fa|sp|su|wi)\s*(\d{2})$/i);
  if (m) {
    const season = { fa: 'fall', sp: 'spring', su: 'summer', wi: 'winter' }[m[1].toLowerCase()];
    mapped = season + '20' + m[2];
  }
  return { target: target, mapped: mapped };
}

function semesterMatches_(rowSemester, activeKeys) {
  const rk = semesterKeys_(rowSemester);
  return rk.mapped === activeKeys.mapped || rk.target === activeKeys.target;
}

// Turns "09_11_2026", "9/11/2026" or "2026-09-11" into 20260911. Only 4-digit years are read,
// so the camera number in "Cam_1_09_11_2026" is never mistaken for part of a date.
function dateKeys_(s) {
  const out = [];
  const str = String(s || '');
  let m;
  const mdy = /(\d{1,2})[_\-\/.](\d{1,2})[_\-\/.](\d{4})/g;
  while ((m = mdy.exec(str))) out.push(+m[3] * 10000 + +m[1] * 100 + +m[2]);
  const ymd = /(\d{4})[_\-\/.](\d{1,2})[_\-\/.](\d{1,2})/g;
  while ((m = ymd.exec(str))) out.push(+m[1] * 10000 + +m[2] * 100 + +m[3]);
  return out;
}

function safeName_(s) { return String(s).replace(/ /g, '_').replace(/,/g, '_'); }

function safeGps_(lat, lon) {
  const f = v => String(v).replace(/\./g, 'd').replace(/-/g, 'n');
  return f(lat) + '_' + f(lon);
}

function cleanStamp_(s) {
  const m = /^\d{4}[:\-\/]\d{2}[:\-\/]\d{2} \d{2}:\d{2}:\d{2}/.exec(String(s || ''));
  return m ? m[0] : '';
}

/* ------------------------------------------------------------------ */
/* Drive folder resolution (runs once per camera, then remembered)     */
/* ------------------------------------------------------------------ */

function listChildFolders_(parent) {
  const out = [];
  const it = parent.getFolders();
  while (it.hasNext()) out.push(it.next());
  return out;
}

function getSemesterFolder_(semString) {
  const k = semesterKeys_(semString);
  const exact = n => { const x = norm_(n); return x === k.target || x === k.mapped; };
  const loose = n => { const x = norm_(n); return x.indexOf(k.target) >= 0 || x.indexOf(k.mapped) >= 0; };
  let candidates = [];
  try {
    const file = DriveApp.getFileById(getSpreadsheet_().getId());
    const parents = file.getParents();
    if (parents.hasNext()) candidates = listChildFolders_(parents.next());
  } catch (e) {}
  let f = candidates.find(x => exact(x.getName())) || candidates.find(x => loose(x.getName()));
  if (f) return f;
  // Last resort (slow): look through every folder in Drive.
  const all = DriveApp.getFolders();
  while (all.hasNext()) {
    const x = all.next();
    if (loose(x.getName())) return x;
  }
  throw new Error('Could not find the semester folder (' + semString + '). Please contact your instructor.');
}

function getSiteFolder_(semester, site) {
  const semFolder = getSemesterFolder_(semester);
  const target = norm_(site);
  const nm = x => norm_(x.getName());
  // Never treat the "<Site> Tagged Photos" folders as the source site folder.
  const list = listChildFolders_(semFolder).filter(x => !/tagged/i.test(x.getName()));
  let f = list.find(x => nm(x) === target)
       || list.find(x => nm(x) === target + 'allphotos')
       || list.find(x => nm(x).indexOf(target) === 0);
  if (!f) {
    // Last resort: any non-tagged folder containing the site name, preferring one that
    // actually holds a CAMERA PHOTOS TO REVIEW folder.
    const loose = list.filter(x => nm(x).indexOf(target) >= 0);
    f = loose.find(x => getTargetSubfolder_(x, 'cameraphotostoreview')) || loose[0];
  }
  if (!f) throw new Error('Site folder not found: ' + site);
  return f;
}

function getTaggedSiteFolder_(semester, site) {
  const semFolder = getSemesterFolder_(semester);
  const target = norm_(site);
  const list = listChildFolders_(semFolder);
  const f = list.find(x => norm_(x.getName()).indexOf(target) >= 0 && /tagged/i.test(x.getName()));
  if (f) return f;
  // Not found — auto-create it so new sites/semesters provision themselves.
  return semFolder.createFolder(site + ' Tagged Photos');
}

function getTargetSubfolder_(parent, nameIncludes) {
  const target = nameIncludes.replace(/ /g, '').toLowerCase();
  const list = listChildFolders_(parent);
  return list.find(f => f.getName().replace(/ /g, '').toLowerCase().indexOf(target) >= 0) || null;
}

function getOrCreateChild_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Picks the deployment folder for a camera. When the same camera number has several
// deployment folders at one site, the folder whose dates match the sheet row wins.
function pickMediaFolder_(reviewBase, cam) {
  const folders = listChildFolders_(reviewBase);
  const num = String(cam.displayCam).trim().replace(/^0+(?=\d)/, '');
  const rx = new RegExp('(?<![a-z])cam[\\s_\\-]*0*' + escapeRx_(num) + '(?![0-9])', 'i');
  let cands = folders.filter(f => rx.test(f.getName()));
  if (!cands.length) {
    const t = ('cam' + cam.displayCam).toLowerCase();
    cands = folders.filter(f => f.getName().replace(/[_ ]/g, '').toLowerCase().indexOf(t) >= 0);
  }
  if (!cands.length) return null;
  if (cands.length === 1) return cands[0];
  const want = dateKeys_(cam.deployment).concat(dateKeys_(cam.collection));
  let best = cands[0], bestScore = -1;
  cands.forEach(f => {
    const have = dateKeys_(f.getName());
    const score = want.filter(w => have.indexOf(w) >= 0).length;
    if (score > bestScore) { best = f; bestScore = score; }
  });
  return best;
}

function folderOk_(id) {
  try {
    const m = Drive.Files.get(id, { fields: 'id,trashed,mimeType', supportsAllDrives: true });
    return !m.trashed && m.mimeType === FOLDER_MIME;
  } catch (e) {
    return false;
  }
}

// Returns { media, mediaName, siteId, tagged? }. IDs are stored in Script Properties so the slow
// folder search only happens the first time a camera is opened. Run clearCaches() after moving
// or re-creating folders.
function getFolderIds_(cam, needTagged) {
  const key = 'f_' + hash_(cam.uid);
  let rec = null;
  try { rec = JSON.parse(props_().getProperty(key) || 'null'); } catch (e) {}
  let changed = false;

  if (rec && rec.media && !folderOk_(rec.media)) rec = null;
  if (!rec || !rec.media) {
    const ov = folderOverrides_(cam);          // folders chosen in the instructor "Camera setup" tab (none = automatic search)
    if (ov.photosCam || ov.photosSite) {
      let mf = null;
      if (ov.photosCam) {
        if (!folderOk_(ov.photosCam)) throw new Error('The photo folder chosen for this camera in Camera setup is missing or in the trash. Please contact your instructor.');
        mf = DriveApp.getFolderById(ov.photosCam);
      } else {
        if (!folderOk_(ov.photosSite)) throw new Error('The photo folder chosen for site ' + cam.site + ' in Camera setup is missing or in the trash. Please contact your instructor.');
        mf = pickMediaFolder_(DriveApp.getFolderById(ov.photosSite), cam);
        if (!mf) throw new Error('Could not locate media folder for Camera ' + cam.displayCam);
      }
      rec = { media: mf.getId(), mediaName: mf.getName(), siteId: '' };
      changed = true;
    } else {
      const siteFolder = getSiteFolder_(cam.semester, cam.site);
      const reviewBase = getTargetSubfolder_(siteFolder, 'cameraphotostoreview');
      if (!reviewBase) throw new Error('Could not find CAMERA PHOTOS TO REVIEW folder in ' + siteFolder.getName());
      const mf = pickMediaFolder_(reviewBase, cam);
      if (!mf) throw new Error('Could not locate media folder for Camera ' + cam.displayCam);
      rec = { media: mf.getId(), mediaName: mf.getName(), siteId: siteFolder.getId() };
      changed = true;
    }
  }

  if (needTagged && (!rec.tagged || !folderOk_(rec.tagged))) {
    const ovT = folderOverrides_(cam);
    if (ovT.taggedCam) {
      if (!folderOk_(ovT.taggedCam)) throw new Error('The tagged-photos folder chosen for this camera in Camera setup is missing or in the trash. Please contact your instructor.');
      rec.tagged = ovT.taggedCam;
    } else {
      let taggedParent;
      if (ovT.taggedSite) {
        if (!folderOk_(ovT.taggedSite)) throw new Error('The tagged-photos folder chosen for site ' + cam.site + ' in Camera setup is missing or in the trash. Please contact your instructor.');
        taggedParent = DriveApp.getFolderById(ovT.taggedSite);
      } else {
        taggedParent = getTaggedSiteFolder_(cam.semester, cam.site);
      }
      const safeUid = cam.uid.replace(/[\/\\]/g, '_');
      const taggedFolder = getOrCreateChild_(taggedParent, safeUid + '_' + cam.site + '_Tagged Photos');
      rec.tagged = taggedFolder.getId();
    }
    changed = true;
  }

  if (changed) props_().setProperty(key, JSON.stringify(rec));
  return rec;
}

// Walks the camera's media folder. Returns the file list (when wanted) and a map of every
// folder ID -> name that belongs to this camera (used to validate later requests).
function walkMedia_(rootId, rootName, taggedId, wantFiles) {
  const tz = Session.getScriptTimeZone();
  const queue = [{ id: rootId, name: rootName }];
  const folders = {};
  folders[rootId] = rootName;
  const out = [];
  let skipped = 0;
  while (queue.length > 0) {
    const cur = queue.shift();
    let pageToken = null;
    do {
      const params = {
        q: "'" + cur.id + "' in parents and trashed = false" + (wantFiles ? '' : " and mimeType = '" + FOLDER_MIME + "'"),
        pageSize: 1000,
        fields: 'nextPageToken, files(id, name, mimeType, modifiedTime, imageMediaMetadata/time, thumbnailLink)',
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      };
      if (pageToken) params.pageToken = pageToken;
      const res = Drive.Files.list(params);
      (res.files || []).forEach(f => {
        if (f.mimeType === FOLDER_MIME) {
          if (f.id === taggedId || /tagged/i.test(f.name)) return;
          folders[f.id] = f.name;
          queue.push({ id: f.id, name: f.name });
          return;
        }
        if (!wantFiles) return;
        if (!CFG.MEDIA_EXT.test(f.name)) { skipped++; return; }
        let time = f.imageMediaMetadata && f.imageMediaMetadata.time;
        if (!time || !/^\d{4}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}/.test(time)) {
          time = Utilities.formatDate(new Date(f.modifiedTime), tz, CFG.STAMP_FMT);
        }
        out.push({
          id: f.id, name: f.name, isVideo: CFG.VIDEO_EXT.test(f.name),
          thumbUrl: f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s1600') : null,
          time: time.slice(0, 19), parentId: cur.id, folder: cur.name,
        });
      });
      pageToken = res.nextPageToken;
    } while (pageToken);
  }
  out.sort((a, b) => (a.time < b.time ? -1 : a.time > b.time ? 1 : a.name.localeCompare(b.name)));
  return { files: out, folders: folders, skipped: skipped };
}

/* ------------------------------------------------------------------ */
/* Sheet access                                                        */
/* ------------------------------------------------------------------ */

function rowToCam_(r, rowNum) {
  const camId = String(r[2] || '').trim();
  const deployment = String(r[3] || '').trim();
  const collection = String(r[4] || '').trim();
  return {
    row: rowNum,
    uid: String(r[11] || '').trim() || (camId + '_' + deployment + '_' + collection),
    site: String(r[0] || '').trim(),
    semester: String(r[1] || '').trim(),
    camId: camId,
    displayCam: camId.replace(/^Cam\s*/i, ''),
    deployment: deployment,
    collection: collection,
    lat: String(r[5] || '').trim(),
    lon: String(r[6] || '').trim(),
    brand: String(r[7] || '').trim(),
    reviewer: String(r[8] || '').trim(),
    complete: String(r[10] || '').trim().toLowerCase() === 'yes',
  };
}

function readDataRows_() {
  const sheet = getSpreadsheet_().getSheetByName('Data');
  const vals = sheet.getDataRange().getDisplayValues();
  const rows = [];
  for (let i = 1; i < vals.length; i++) {
    const c = rowToCam_(vals[i], i + 1);
    if (c.site && c.camId) rows.push(c);
  }
  return { sheet: sheet, rows: rows };
}

function activeSemesterKeys_() {
  const studentSheet = getSpreadsheet_().getSheetByName('Student Names');
  return semesterKeys_(String(studentSheet.getRange('A1').getValue()).trim());
}

function publicCam_(c) {
  const dateRange = (c.deployment && c.collection) ? ' ' + c.deployment + ' to ' + c.collection
    : (c.deployment || c.collection) ? ' ' + (c.deployment || c.collection) : '';
  return {
    key: c.uid,
    label: 'Cam ' + c.displayCam + ' (' + c.site + ')' + dateRange,
    site: c.site, semester: c.semester, camId: c.camId, uniqueCamId: c.uid,
    displayCam: c.displayCam, deployment: c.deployment, collection: c.collection,
    lat: c.lat, lon: c.lon, brand: c.brand, reviewer: c.reviewer,
  };
}

/* ------------------------------------------------------------------ */
/* Species list                                                        */
/* ------------------------------------------------------------------ */

function getSpeciesFileId_() {
  const p = props_();
  const saved = p.getProperty('SPECIES_FILE_ID');
  if (saved) return saved;
  const it = DriveApp.getFilesByName('species');
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      p.setProperty('SPECIES_FILE_ID', f.getId());
      return f.getId();
    }
  }
  return null;
}

function speciesSheet_() {
  for (let attempt = 0; attempt < 2; attempt++) {
    const id = getSpeciesFileId_();
    if (!id) break;
    try { return SpreadsheetApp.openById(id).getSheets()[0]; }
    catch (e) { props_().deleteProperty('SPECIES_FILE_ID'); }
  }
  throw new Error('species sheet not found');
}

function readSpeciesSheet_() {
  const sh = speciesSheet_();
  const last = sh.getLastRow();
  if (last < 1) return [];
  return sh.getRange(1, 1, last, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function getSpecies_() {
  const c = cache_();
  const hit = c.get('species');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  let list;
  try { list = readSpeciesSheet_(); } catch (e) { list = DEFAULT_SPECIES; }
  try { c.put('species', JSON.stringify(list), 600); } catch (e) {}
  return list;
}

function validSpecies_(s) {
  return typeof s === 'string' && s.length > 0 && s.length <= 80 && !/^[=+\-@|]/.test(s);
}

// Returns a Map of lower-case name -> canonical name. Unknown names are added to the species
// sheet under a lock so two students adding the same species cannot create duplicates.
function canonicalizeSpecies_(names) {
  const result = new Map();
  if (!names.length) return result;
  const build = list => {
    const m = new Map();
    list.filter(s => s.indexOf('|||') !== 0).forEach(s => m.set(s.toLowerCase(), s));
    return m;
  };
  let lookup = build(getSpecies_());
  const missing = new Map();
  names.forEach(n => { const k = n.toLowerCase(); if (!lookup.has(k) && !missing.has(k)) missing.set(k, n); });
  if (missing.size) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      let fresh;
      try { fresh = readSpeciesSheet_(); } catch (e) { fresh = getSpecies_(); }
      lookup = build(fresh);
      const toAdd = [];
      missing.forEach((n, k) => { if (!lookup.has(k)) { lookup.set(k, n); toAdd.push(n); } });
      if (toAdd.length) {
        try {
          const sh = speciesSheet_();
          sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 1).setValues(toAdd.map(s => [s]));
        } catch (e) {}
        cache_().remove('species');
      }
    } finally {
      lock.releaseLock();
    }
  }
  names.forEach(n => result.set(n.toLowerCase(), lookup.get(n.toLowerCase()) || n));
  return result;
}

/* ------------------------------------------------------------------ */
/* Login, sessions, claims                                             */
/* ------------------------------------------------------------------ */

function checkCredentials_(nameStr, pinStr) {
  const name = String(nameStr || '').trim();
  const pin = String(pinStr || '').trim();
  const cache = cache_();
  const failKey = 'fail_' + hash_(name.toLowerCase());
  const fails = parseInt(cache.get(failKey) || '0', 10);
  if (fails >= CFG.MAX_FAILS) {
    throw new Error('Too many failed attempts. Please wait 15 minutes and try again.');
  }
  const vals = getSpreadsheet_().getSheetByName('Student Names').getDataRange().getDisplayValues();
  for (let i = 1; i < vals.length; i++) {
    const sName = String(vals[i][0]).trim();
    if (!sName) continue;
    const sPin = String(vals[i][1]).trim().padStart(4, '0');
    if (sName.toLowerCase() === name.toLowerCase() && sPin === pin) {
      cache.remove(failKey);
      return sName;
    }
  }
  cache.put(failKey, String(fails + 1), CFG.FAIL_WINDOW);
  throw new Error('Invalid Name or PIN. Please check your credentials.');
}

function newSession_(data) {
  const id = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache_().put('sess_' + id, JSON.stringify(data), CFG.SESSION_TTL);
  return id;
}

function loadSession_(sess) {
  if (!sess) throw new Error('SESSION_EXPIRED');
  const c = cache_();
  const raw = c.get('sess_' + sess);
  if (!raw) throw new Error('SESSION_EXPIRED');
  const S = JSON.parse(raw);
  if (Date.now() - S.t > 3600000) {            // slide the expiry forward at most once an hour
    S.t = Date.now();
    c.put('sess_' + sess, JSON.stringify(S), CFG.SESSION_TTL);
  }
  return S;
}

// Session of a student who is reviewing a camera.
function requireSession_(sess) {
  const S = loadSession_(sess);
  if (S.viewer) throw new Error('Not allowed.');
  return S;
}

// Session of a student who signed in to browse tagged photos (no camera attached).
function requireViewer_(sess) {
  const S = loadSession_(sess);
  if (!S.viewer) throw new Error('SESSION_EXPIRED');
  return S;
}

// Run once from the Apps Script editor (Run > setupInstructorColumn): adds an "Instructor" column to the
// Student Names tab if there is none, and puts "yes" beside names in CFG.ADMIN_NAMES. Nothing else is touched.
function setupInstructorColumn() {
  const sh = getSpreadsheet_().getSheetByName('Student Names');
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let col = head.findIndex(h => String(h).trim().toLowerCase() === 'instructor') + 1;
  if (!col) { col = lastCol + 1; sh.getRange(1, col).setValue('Instructor'); }
  const last = sh.getLastRow();
  if (last >= 2) {
    const names = sh.getRange(2, 1, last - 1, 1).getValues();
    const cur = sh.getRange(2, col, last - 1, 1).getValues();
    names.forEach((r, i) => {
      const n = String(r[0]).trim().toLowerCase();
      if (!cur[i][0] && n && CFG.ADMIN_NAMES.some(a => n === a || n.indexOf(a + ' ') === 0)) sh.getRange(i + 2, col).setValue('yes');
    });
  }
  return 'Instructor column is column ' + col;
}

// Instructor = "yes" in an "Instructor" column (header in row 1) of the Student Names tab.
// The names in CFG.ADMIN_NAMES remain instructors too, so existing access keeps working.
function isAdmin_(name) {
  const n = String(name || '').trim().toLowerCase();
  if (CFG.ADMIN_NAMES.some(a => n === a || n.indexOf(a + ' ') === 0)) return true;
  if (!n) return false;
  try {
    const vals = getSpreadsheet_().getSheetByName('Student Names').getDataRange().getDisplayValues();
    if (!vals.length) return false;
    let col = -1;
    for (let c = 0; c < vals[0].length; c++) if (String(vals[0][c]).trim().toLowerCase() === 'instructor') { col = c; break; }
    if (col < 0) return false;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim().toLowerCase() === n) return /^(yes|y|true|x)$/i.test(String(vals[i][col]).trim());
    }
  } catch (e) {}
  return false;
}

function makeSession_(name, cam, ids, folders) {
  return newSession_({
    name: name,
    mediaId: ids.media,
    taggedId: ids.tagged,
    folders: folders,
    cam: {
      uid: cam.uid, site: cam.site, semester: cam.semester, camId: cam.camId,
      displayCam: cam.displayCam, lat: cam.lat, lon: cam.lon,
    },
    t: Date.now(),
  });
}

function publicConfig_(cam, name) {
  return {
    key: cam.uid, site: cam.site, semester: cam.semester, camId: cam.camId,
    uniqueCamId: cam.uid, displayCam: cam.displayCam, deployment: cam.deployment,
    collection: cam.collection, lat: cam.lat, lon: cam.lon, brand: cam.brand,
    reviewer: name, label: 'Cam ' + cam.displayCam + ' (' + cam.site + ')',
  };
}

// Finds the student's active camera, or claims the requested one. Runs under a lock so two
// students cannot claim the same camera.
function claimCamera_(realName, uid) {
  const keys = activeSemesterKeys_();
  const lname = realName.toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = readDataRows_();
    let target = data.rows.find(c => c.reviewer.toLowerCase() === lname && !c.complete && semesterMatches_(c.semester, keys));
    if (!target) {
      if (!uid) throw new Error('Please select a camera to review.');
      target = data.rows.find(c => c.uid === uid);
      if (!target) throw new Error('That camera is no longer on the list. Refresh the page and try again.');
      if (target.complete) throw new Error('That review is already complete.');
      if (target.reviewer && target.reviewer.toLowerCase() !== lname) {
        throw new Error('This camera is already claimed by ' + target.reviewer);
      }
      data.sheet.getRange(target.row, 9).setValue(realName);
      target.reviewer = realName;
    }
    return target;
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Progress                                                            */
/* ------------------------------------------------------------------ */

function progKey_(uid, reviewer) { return 'p_' + hash_(uid + '|' + String(reviewer).trim().toLowerCase()); }
// Older versions keyed progress on the sheet row number; still read it so nobody loses their place.
function legacyProgKey_(row, reviewer) { return 'p_' + hash_(row + '|' + String(reviewer).trim().toLowerCase()); }

function startIndexFor_(cam, name, files) {
  const p = props_();
  const saved = p.getProperty(progKey_(cam.uid, name)) || p.getProperty(legacyProgKey_(cam.row, name));
  if (!saved || !files.length) return 0;
  try {
    const s = JSON.parse(saved);
    const i = files.findIndex(f => f.id === s.id);
    return i >= 0 ? i : Math.min(s.idx || 0, files.length - 1);
  } catch (e) {
    return 0;
  }
}

function writeProgress_(S, fileId, idx) {
  props_().setProperty(progKey_(S.cam.uid, S.name), JSON.stringify({ id: String(fileId || ''), idx: parseInt(idx, 10) || 0 }));
}

/* ------------------------------------------------------------------ */
/* Functions the browser calls - before login                          */
/* ------------------------------------------------------------------ */

function getCameras() {
  const keys = activeSemesterKeys_();
  const data = readDataRows_();
  const sites = new Set();
  const out = [];
  data.rows.forEach(c => {
    if (!semesterMatches_(c.semester, keys)) return;
    sites.add(c.site);
    if (c.complete) return;
    out.push(publicCam_(c));
  });
  out.sort((a, b) =>
    a.site.localeCompare(b.site) ||
    String(a.displayCam).localeCompare(String(b.displayCam), undefined, { numeric: true }) ||
    a.deployment.localeCompare(b.deployment, undefined, { numeric: true }));
  return { cameras: out, sites: Array.from(sites).sort() };
}

function getBootstrap() {
  const studentSheet = getSpreadsheet_().getSheetByName('Student Names');
  const last = studentSheet.getLastRow();
  const students = [];
  if (last >= 2) {
    studentSheet.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(r => {
      const n = String(r[0]).trim();
      if (n) students.push(n);
    });
  }
  const cams = getCameras();
  return { cameras: cams.cameras, sites: cams.sites, species: getSpecies_(), students: students };
}

function getPreviewImage(uid) {
  const data = readDataRows_();
  const cam = data.rows.find(c => c.uid === String(uid));
  if (!cam) throw new Error('Camera not found. Refresh the page and try again.');
  const ids = getFolderIds_(cam, false);
  const queue = [ids.media];
  let guard = 0, seenFiles = 0, seenImages = 0, fallback = null;
  while (queue.length > 0 && guard++ < 30) {
    const cur = queue.shift();
    const res = Drive.Files.list({
      q: "'" + cur + "' in parents and trashed = false",
      pageSize: 100,
      orderBy: 'name',
      fields: 'files(id, name, mimeType, size, thumbnailLink)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    const files = res.files || [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      if (f.mimeType === FOLDER_MIME) {
        if (!/tagged/i.test(f.name || '')) queue.push(f.id);
        continue;
      }
      seenFiles++;
      const isImage = String(f.mimeType || '').indexOf('image/') === 0 || /\.(png|jpe?g)$/i.test(f.name || '');
      if (!isImage) continue;
      seenImages++;
      if (f.thumbnailLink) return f.thumbnailLink.replace(/=s\d+$/, '=s800');
      if (!fallback) fallback = f;   // image without a Drive thumbnail (e.g. just uploaded)
    }
  }
  // Drive has not generated thumbnails yet (common right after a bulk upload): send the image itself.
  if (fallback) {
    try {
      const file = DriveApp.getFileById(fallback.id);
      const blob = file.getBlob();
      if (blob.getBytes().length <= 8 * 1024 * 1024) {
        return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
    } catch (e) {}
  }
  throw new Error('No preview is available for this camera. Looked in "' + ids.mediaName + '": found ' +
    seenFiles + ' file(s), ' + seenImages + ' image(s). ' +
    (seenFiles === 0 ? 'The folder looks empty - if the photos were just uploaded, wait for Drive to finish syncing, or check that this is the right Cam folder.' : ''));
}

function authenticateAndLoad(nameStr, pinStr, uid) {
  const realName = checkCredentials_(nameStr, pinStr);
  const cam = claimCamera_(realName, uid ? String(uid) : '');
  const ids = getFolderIds_(cam, true);
  const tree = walkMedia_(ids.media, ids.mediaName, ids.tagged, true);
  const sess = makeSession_(realName, cam, ids, tree.folders);
  return {
    sess: sess,
    token: ScriptApp.getOAuthToken(),
    config: publicConfig_(cam, realName),
    files: tree.files,
    startIndex: startIndexFor_(cam, realName, tree.files),
    skipped: tree.skipped,
  };
}

// Used by the browser to get a fresh session after the old one expired (idle > 6 hours).
function relogin(nameStr, pinStr) {
  const realName = checkCredentials_(nameStr, pinStr);
  const cam = claimCamera_(realName, '');
  const ids = getFolderIds_(cam, true);
  const tree = walkMedia_(ids.media, ids.mediaName, ids.tagged, false);
  return { sess: makeSession_(realName, cam, ids, tree.folders), token: ScriptApp.getOAuthToken() };
}

/* ------------------------------------------------------------------ */
/* Functions the browser calls - after login (first argument: session) */
/* ------------------------------------------------------------------ */

function refreshToken(sess) {
  requireSession_(sess);
  return ScriptApp.getOAuthToken();
}

function saveProgress(sess, fileId, idx) {
  const S = requireSession_(sess);
  writeProgress_(S, fileId, idx);
  return true;
}

function getCameraTaggedPhotos(sess) {
  const S = requireSession_(sess);
  if (!S.taggedId) return [];                        // instructor session on a camera with no tagged folder yet
  const info = taggedInfoMap_(S.cam);
  const photos = [];
  let pageToken = null;
  do {
    const params = {
      q: "'" + S.taggedId + "' in parents and trashed = false",
      pageSize: 500,
      fields: 'nextPageToken, files(id, name, mimeType, thumbnailLink)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    };
    if (pageToken) params.pageToken = pageToken;
    const res = Drive.Files.list(params);
    (res.files || []).forEach(f => {
      if (f.mimeType.indexOf('image/') === 0 || f.mimeType.indexOf('video/') === 0) {
        const rec = info.byId[f.id] || info.byName[f.name] || null;
        photos.push({
          id: f.id,
          name: f.name,
          species: rec ? rec.species : [],
          thumb: f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s300') : '',
          big: f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s1200') : '',
          isVideo: f.mimeType.indexOf('video/') === 0,
        });
      }
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  photos.sort((a, b) => a.name.localeCompare(b.name));
  return photos;
}

function masterSheet_(createIfMissing) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName('Master Data');
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet('Master Data');
    sheet.getRange(1, 1, 1, CFG.MASTER_HEADERS.length).setValues([CFG.MASTER_HEADERS]);
  }
  return sheet;
}

// Makes sure the Master Data sheet has every column (File_ID, Stars, ID_Doubt, Suggested_ID).
// Existing rows get 0 stars / 0 doubts. Call while holding the script lock.
function ensureMasterColumns_(sheet) {
  const total = CFG.MASTER_HEADERS.length;
  if (sheet.getMaxColumns() < total) sheet.insertColumnsAfter(sheet.getMaxColumns(), total - sheet.getMaxColumns());
  const hdr = sheet.getRange(1, 1, 1, total).getValues()[0];
  if (!String(hdr[CFG.FILE_ID_COL - 1]).trim()) sheet.getRange(1, CFG.FILE_ID_COL).setValue('File_ID');
  const last = sheet.getLastRow();
  for (let c = CFG.STARS_COL; c <= total; c++) {
    if (String(hdr[c - 1]).trim()) continue;
    sheet.getRange(1, c).setValue(CFG.MASTER_HEADERS[c - 1]);
    if (last >= 2 && c !== CFG.SUGG_COL) sheet.getRange(2, c, last - 1, 1).setValue(0);
  }
}

function appendMasterRows_(rows) {
  if (!rows.length) return;
  withRetry_(() => {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sheet = masterSheet_(true);
      const total = CFG.MASTER_HEADERS.length;
      ensureMasterColumns_(sheet);
      const start = sheet.getLastRow() + 1;
      const need = start + rows.length - 1;
      if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
      const padded = rows.map(r => {
        const x = r.slice();
        while (x.length < total) x.push(x.length === CFG.SUGG_COL - 1 ? '' : (x.length >= CFG.STARS_COL - 1 ? 0 : ''));
        return x;
      });
      sheet.getRange(start, 1, padded.length, total).setValues(padded);
    } finally {
      lock.releaseLock();
    }
  }, 2);
}

function masterHasFileId_(fileId) {
  const sheet = masterSheet_(false);
  const nCols = CFG.FILE_ID_COL;
  if (!sheet || sheet.getLastRow() < 2 || sheet.getMaxColumns() < nCols) return false;
  return !!sheet.getRange(2, nCols, sheet.getLastRow() - 1, 1).createTextFinder(fileId).matchEntireCell(true).findNext();
}

function removeMasterRow_(fileId, newName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = masterSheet_(false);
    if (!sheet || sheet.getLastRow() < 2) return;
    const nCols = CFG.FILE_ID_COL;
    let row = 0;
    if (sheet.getMaxColumns() >= nCols) {
      const hit = sheet.getRange(2, nCols, sheet.getLastRow() - 1, 1).createTextFinder(fileId).matchEntireCell(true).findNext();
      if (hit) row = hit.getRow();
    }
    if (!row && newName) {                      // rows written by older versions have no File_ID
      const names = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
      for (let i = names.length - 1; i >= 0; i--) {
        if (names[i][0] === newName) { row = i + 2; break; }
      }
    }
    if (row) sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
}

// Tags a batch of files. Each file is moved and renamed with a single Drive call.
// Safe to retry: a file that is already in the tagged folder is treated as done.
function tagFiles(sess, tasks, progress) {
  const S = requireSession_(sess);
  if (!S.taggedId) ensureTaggedFolder_(sess, S);     // instructor sessions only: students always have one
  const cam = S.cam;
  const cleaned = (Array.isArray(tasks) ? tasks : []).slice(0, CFG.MAX_BATCH).map(t => ({
    fileId: String((t && t.fileId) || ''),
    species: (Array.isArray(t && t.species) ? t.species : []).map(s => String(s).trim().replace(/\s+/g, ' ')).filter(Boolean),
    originalTime: cleanStamp_(t && t.originalTime),
    correctedTime: cleanStamp_(t && t.correctedTime),
  }));
  const ok = cleaned.map(t => !!t.fileId && t.species.length > 0 && t.species.every(validSpecies_));
  const allNames = [];
  cleaned.forEach((t, i) => { if (ok[i]) t.species.forEach(s => allNames.push(s)); });
  const canon = canonicalizeSpecies_(allNames);

  const results = [];
  const logRows = [];
  cleaned.forEach((t, i) => {
    const res = { fileId: t.fileId };
    results.push(res);
    if (!ok[i]) { res.error = 'Invalid species name or file.'; return; }
    try {
      const names = t.species.map(s => canon.get(s.toLowerCase()) || s);
      let meta;
      try {
        meta = Drive.Files.get(t.fileId, { fields: 'id,name,parents,trashed', supportsAllDrives: true });
      } catch (e) { res.error = 'missing'; return; }
      const parent = (meta.parents || [])[0];
      if (meta.trashed) { res.error = 'missing'; return; }

      if (parent === S.taggedId) {
        // Already moved by an earlier attempt: make sure it is logged, then report success.
        res.ok = true; res.noUndo = true; res.newFileId = t.fileId; res.newName = meta.name; res.originalName = meta.name;
        if (!masterHasFileId_(t.fileId)) {
          logRows.push(buildLogRow_(S, cam, t, meta.name, names, ''));
        }
        return;
      }
      if (!parent || !Object.prototype.hasOwnProperty.call(S.folders, parent)) { res.error = 'missing'; return; }

      const m = /^(\d{4})[:\-\/](\d{2})[:\-\/](\d{2}) (\d{2}):(\d{2}):(\d{2})/.exec(t.correctedTime);
      const dateStr = m ? m[1] + m[2] + m[3] : 'UnknownDate';
      const timeStr = m ? m[4] + m[5] + m[6] : 'UnknownTime';
      const origFull = meta.name;
      const dot = origFull.lastIndexOf('.');
      const origName = dot > 0 ? origFull.slice(0, dot) : origFull;
      const ext = dot > 0 ? origFull.slice(dot) : '';
      const newName = buildTaggedName_(cam, names, origName, ext);

      Drive.Files.update({ name: newName }, t.fileId, null, {
        addParents: S.taggedId, removeParents: parent, supportsAllDrives: true,
      });

      res.ok = true;
      res.tagId = Utilities.getUuid();
      res.newFileId = t.fileId;
      res.newName = newName;
      res.originalName = origFull;
      res.originalParentId = parent;
      logRows.push(buildLogRow_(S, cam, t, newName, names, S.folders[parent]));
    } catch (e) {
      res.error = String(e && e.message ? e.message : e);
    }
  });

  let logWarning = '';
  try {
    appendMasterRows_(logRows);
  } catch (e) {
    logWarning = 'Files were tagged, but the Master Data log could not be updated (' + e.message + '). Tell your instructor.';
  }
  if (progress && progress.fileId !== undefined) {
    try { writeProgress_(S, progress.fileId, progress.idx); } catch (e) {}
  }
  return { results: results, logWarning: logWarning, species: getSpecies_() };
}

function buildLogRow_(S, cam, t, newName, names, parentName) {
  return [
    'https://drive.google.com/file/d/' + t.fileId + '/view',
    parentName,
    S.name,
    cam.site,
    cam.uid,
    cam.camId,
    newName,
    names.join(', '),
    cam.lat + ', ' + cam.lon,
    t.originalTime,
    t.correctedTime,
    t.correctedTime === t.originalTime ? 'Correct' : 'Incorrect',
    '0',
    t.fileId,
  ];
}

/* ------------------------------------------------------------------ */
/* Fixing old tags: change a species / remove a tag                    */
/* ------------------------------------------------------------------ */

// Name given to a tagged file: <site>_<species>_<camera id>_<gps>_<original name><ext>
function buildTaggedName_(cam, names, origName, ext) {
  const safeSpecies = names.map(s => s.replace(/ /g, '_').replace(/\./g, '')).join('_');
  return [
    safeName_(cam.site), safeSpecies, safeName_(cam.uid.replace(/[\/\\]/g, '_')),
    safeGps_(cam.lat, cam.lon), origName
  ].join('_') + ext;
}

// Recovers the camera's original file name (with extension) from a tagged file name.
// Returns '' when the name does not follow the expected pattern.
function originalNameFromTagged_(taggedName, cam) {
  const uidSeg = '_' + safeName_(cam.uid.replace(/[\/\\]/g, '_')) + '_';
  const at = String(taggedName).indexOf(uidSeg);
  if (at < 0) return '';
  const parts = String(taggedName).slice(at + uidSeg.length).split('_');   // lat, lon, original name...
  if (parts.length < 3) return '';
  return parts.slice(2).join('_');
}

function stampText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), CFG.STAMP_FMT);
  return String(v || '');
}

// Master Data rows for one camera, keyed by Drive file ID and by file name (older rows have no ID).
function taggedInfoMap_(cam) {
  const out = { byId: {}, byName: {} };
  const sheet = masterSheet_(false);
  if (!sheet || sheet.getLastRow() < 2) return out;
  const nCols = CFG.FILE_ID_COL;
  const width = Math.min(nCols, sheet.getMaxColumns());
  if (width < 8) return out;
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  vals.forEach(r => {
    if (String(r[4]) !== cam.uid) return;
    const rec = {
      species: String(r[7] || '').split(',').map(x => x.trim()).filter(Boolean),
      folder: String(r[1] || ''),
      time: width > 9 ? stampText_(r[9]) : '',
    };
    if (width >= nCols && r[13]) out.byId[String(r[13])] = rec;
    if (r[6]) out.byName[String(r[6])] = rec;
  });
  return out;
}

function findMasterRow_(sheet, fileId, fileName) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const nCols = CFG.FILE_ID_COL;
  let row = 0;
  if (sheet.getMaxColumns() >= nCols) {
    const hit = sheet.getRange(2, nCols, sheet.getLastRow() - 1, 1).createTextFinder(fileId).matchEntireCell(true).findNext();
    if (hit) row = hit.getRow();
  }
  if (!row && fileName) {                     // rows written by older versions have no File_ID
    const names = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
    for (let i = names.length - 1; i >= 0; i--) {
      if (names[i][0] === fileName) { row = i + 2; break; }
    }
  }
  return row;
}

// Updates the File_Name and Species cells of a photo's Master Data row (or logs a row if it has none).
function updateMasterTag_(S, fileId, oldName, newName, names) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = masterSheet_(true);
    const row = findMasterRow_(sheet, fileId, oldName);
    if (row) {
      sheet.getRange(row, 7, 1, 2).setValues([[newName, names.join(', ')]]);
      const nCols = CFG.FILE_ID_COL;
      if (sheet.getMaxColumns() >= nCols && !String(sheet.getRange(row, nCols).getValue()).trim()) {
        sheet.getRange(row, nCols).setValue(fileId);
      }
      return;
    }
  } finally {
    lock.releaseLock();
  }
  const rowData = buildLogRow_(S, S.cam, { fileId: fileId, originalTime: '', correctedTime: '' }, newName, names, '');
  rowData[11] = 'Unknown';
  appendMasterRows_([rowData]);
}

// Reads a file that is (or was) part of this camera and returns it in the shape the review screen uses.
function mediaRecordFor_(fileId, parentId, parentName, savedTime) {
  const f = Drive.Files.get(fileId, {
    fields: 'id,name,mimeType,modifiedTime,thumbnailLink,imageMediaMetadata/time',
    supportsAllDrives: true,
  });
  let time = cleanStamp_(savedTime);
  if (!time) {
    const t = f.imageMediaMetadata && f.imageMediaMetadata.time;
    time = (t && /^\d{4}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}/.test(t))
      ? t : Utilities.formatDate(new Date(f.modifiedTime), Session.getScriptTimeZone(), CFG.STAMP_FMT);
  }
  return {
    id: f.id, name: f.name, isVideo: CFG.VIDEO_EXT.test(f.name),
    thumbUrl: f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s1600') : null,
    time: time.slice(0, 19), parentId: parentId, folder: parentName,
  };
}

// Picks the folder a tagged photo came from, using the folder name saved in Master Data.
function homeFolderFor_(S, folderName) {
  const ids = Object.keys(S.folders).filter(id => S.folders[id] === folderName);
  if (ids.length === 1) return ids[0];
  if (ids.indexOf(S.mediaId) >= 0) return S.mediaId;
  return ids[0] || S.mediaId;
}

// Shared by the reviewer ("Tagged photos" view) and the instructor ("View tagged photos" window).
// ctx = { cam, taggedId, name }. Renames the file, updates Master Data, and clears the "Nahhh" flags
// and suggested IDs on the photo (the question they raised has been answered). Stars are kept.
function changeTagCore_(ctx, fileId, species) {
  fileId = String(fileId || '');
  const seen = {};
  const cleaned = (Array.isArray(species) ? species : [])
    .map(s => String(s).trim().replace(/\s+/g, ' ')).filter(Boolean)
    .filter(s => { const k = s.toLowerCase(); if (seen[k]) return false; seen[k] = true; return true; });
  if (!fileId || !cleaned.length || !cleaned.every(validSpecies_)) {
    return { error: 'Please enter at least one valid species name.' };
  }
  try {
    let meta;
    try {
      meta = Drive.Files.get(fileId, { fields: 'id,name,parents,trashed', supportsAllDrives: true });
    } catch (e) { return { error: 'That photo could not be found.' }; }
    if (meta.trashed || !ctx.taggedId || (meta.parents || [])[0] !== ctx.taggedId) {
      return { error: "That photo is not in this camera's tagged folder." };
    }
    const origFull = originalNameFromTagged_(meta.name, ctx.cam);
    if (!origFull) {
      return { error: "Could not read this photo's original file name from \"" + meta.name + '". Ask your instructor to fix it.' };
    }
    const canon = canonicalizeSpecies_(cleaned);
    const names = cleaned.map(s => canon.get(s.toLowerCase()) || s);
    const dot = origFull.lastIndexOf('.');
    const origName = dot > 0 ? origFull.slice(0, dot) : origFull;
    const ext = dot > 0 ? origFull.slice(dot) : '';
    const newName = buildTaggedName_(ctx.cam, names, origName, ext);

    if (newName !== meta.name) {
      Drive.Files.update({ name: newName }, fileId, null, { supportsAllDrives: true });
    }
    let warning = '';
    try {
      updateMasterTag_(ctx, fileId, meta.name, newName, names);
    } catch (e) {
      warning = 'The photo was renamed, but the Master Data log could not be updated (' + e.message + '). Tell your instructor.';
    }
    try {
      clearDoubt_(fileId, newName);
    } catch (e) {
      warning = warning || 'The photo was renamed, but its Nahhh / suggestion marks could not be cleared (' + e.message + ').';
    }
    return { ok: true, name: newName, species: names, warning: warning, speciesList: getSpecies_() };
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  }
}

// Changes the species on a photo that is already tagged: renames the file and updates Master Data.
function changeTaggedSpecies(sess, fileId, species) {
  const S = requireSession_(sess);
  return changeTagCore_({ cam: S.cam, taggedId: S.taggedId, name: S.name }, fileId, species);
}

// Removes a tag: the photo goes back to its original folder with its original name (so it shows up
// in the review queue again) and its Master Data row is deleted. Safe to retry.
function removeTaggedPhoto(sess, fileId) {
  const S = requireSession_(sess);
  fileId = String(fileId || '');
  if (!fileId) return { error: 'No photo selected.' };
  if (!S.taggedId) return { error: 'There are no tagged photos for this camera yet.' };
  try {
    let meta;
    try {
      meta = Drive.Files.get(fileId, { fields: 'id,name,parents,trashed', supportsAllDrives: true });
    } catch (e) { return { error: 'That photo could not be found.' }; }
    if (meta.trashed) return { error: 'That photo could not be found.' };
    const parent = (meta.parents || [])[0];
    const info = taggedInfoMap_(S.cam);
    const rec = info.byId[fileId] || info.byName[meta.name] || null;

    if (parent !== S.taggedId) {
      // Already moved back by an earlier attempt: just make sure the log row is gone.
      if (parent && Object.prototype.hasOwnProperty.call(S.folders, parent)) {
        try { removeMasterRow_(fileId, ''); deleteFeedbackFor_(fileId); } catch (e) {}
        return { ok: true, file: mediaRecordFor_(fileId, parent, S.folders[parent], rec && rec.time), warning: '' };
      }
      return { error: "That photo is not in this camera's tagged folder." };
    }

    const origFull = originalNameFromTagged_(meta.name, S.cam);
    if (!origFull) {
      return { error: "Could not work out this photo's original file name from \"" + meta.name + '", so it cannot be returned to the review queue. Ask your instructor.' };
    }
    const homeId = homeFolderFor_(S, rec ? rec.folder : '');
    Drive.Files.update({ name: origFull }, fileId, null, {
      addParents: homeId, removeParents: S.taggedId, supportsAllDrives: true,
    });
    let warning = '';
    try {
      removeMasterRow_(fileId, meta.name);
    } catch (e) {
      warning = 'The photo was returned to the review queue, but its Master Data row could not be removed (' + e.message + '). Tell your instructor.';
    }
    try { deleteFeedbackFor_(fileId); } catch (e) {}
    return { ok: true, file: mediaRecordFor_(fileId, homeId, S.folders[homeId] || '', rec && rec.time), warning: warning };
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  }
}

function undoTag(sess, u) {
  const S = requireSession_(sess);
  try {
    const fileId = String(u && u.newFileId || '');
    const homeId = String(u && u.originalParentId || '');
    if (!fileId || !Object.prototype.hasOwnProperty.call(S.folders, homeId)) {
      throw new Error('That file no longer belongs to this camera.');
    }
    const meta = Drive.Files.get(fileId, { fields: 'id,parents,trashed', supportsAllDrives: true });
    if (meta.trashed || (meta.parents || [])[0] !== S.taggedId) {
      throw new Error('The tagged file could not be found in the tagged folder.');
    }
    Drive.Files.update({ name: String(u.originalName) }, fileId, null, {
      addParents: homeId, removeParents: S.taggedId, supportsAllDrives: true,
    });
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  }
  let warning = '';
  try {
    removeMasterRow_(String(u.newFileId), String(u.newName || ''));
  } catch (e) {
    warning = 'File restored, but the Master Data row could not be removed (' + e.message + '). Tell your instructor.';
  }
  try { deleteFeedbackFor_(String(u.newFileId)); } catch (e) {}
  return { ok: true, warning: warning };
}

function markReviewComplete(sess) {
  const S = requireSession_(sess);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = readDataRows_();
    const cam = data.rows.find(c => c.uid === S.cam.uid);
    if (!cam) throw new Error('Camera not found.');
    if (cam.reviewer.toLowerCase() !== S.name.toLowerCase()) {
      throw new Error('You are not the assigned reviewer for this camera.');
    }
    data.sheet.getRange(cam.row, 11).setValue('Yes');
  } finally {
    lock.releaseLock();
  }
  props_().deleteProperty(progKey_(S.cam.uid, S.name));
  cache_().remove('sess_' + sess);
  return true;
}

/* ------------------------------------------------------------------ */
/* View tagged photos: browse, star, "Nahhh", suggest, instructor fix   */
/* ------------------------------------------------------------------ */

function feedbackSheet_(createIfMissing) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CFG.FEEDBACK_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(CFG.FEEDBACK_SHEET);
    sh.getRange(1, 1, 1, CFG.FEEDBACK_HEADERS.length).setValues([CFG.FEEDBACK_HEADERS]);
    sh.getRange(1, 1, sh.getMaxRows(), 3).setNumberFormat('@');   // IDs and names stay text
    sh.setFrozenRows(1);
  }
  return sh;
}

function splitSugg_(v) {
  return String(v || '').split('|').map(x => x.trim()).filter(Boolean);
}

// All feedback rows, grouped by Drive file ID.
function readFeedbackAll_() {
  const out = { byId: {} };
  const sh = feedbackSheet_(false);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(r => {
    const id = String(r[0] || '');
    if (!id) return;
    (out.byId[id] = out.byId[id] || []).push({
      student: String(r[2] || ''), star: Number(r[3]) === 1, nahh: Number(r[4]) === 1, sugg: splitSugg_(r[5]),
    });
  });
  return out;
}

// Totals for one photo plus what the signed-in student did. Admins also get the list of who did what.
function summarizeFeedback_(rows, me, withVotes) {
  const counts = {}, order = [];
  let stars = 0, nahhs = 0, myStar = false, myNahh = false, mySugg = [];
  const meKey = String(me || '').toLowerCase();
  (rows || []).forEach(r => {
    if (r.star) stars++;
    if (r.nahh) nahhs++;
    r.sugg.forEach(x => {
      const k = x.toLowerCase();
      if (!counts[k]) { counts[k] = { name: x, n: 0 }; order.push(k); }
      counts[k].n++;
    });
    if (r.student.toLowerCase() === meKey) { myStar = r.star; myNahh = r.nahh; mySugg = r.sugg.slice(); }
  });
  const sugg = order.map(k => counts[k]).sort((a, b) => b.n - a.n || a.name.localeCompare(b.name));
  const out = { stars: stars, nahhs: nahhs, sugg: sugg, myStar: myStar, myNahh: myNahh, mySugg: mySugg };
  if (withVotes) out.votes = (rows || []).map(r => ({ student: r.student, star: r.star, nahh: r.nahh, sugg: r.sugg }));
  return out;
}

// Writes Stars / ID_Doubt / Suggested_ID into the photo's Master Data row. Call while holding the lock.
function writeMasterFeedback_(fileId, fileName, sum) {
  const sheet = masterSheet_(false);
  if (!sheet) return false;
  ensureMasterColumns_(sheet);
  const row = findMasterRow_(sheet, fileId, fileName);
  if (!row) return false;
  sheet.getRange(row, CFG.STARS_COL, 1, 3).setValues([[sum.stars, sum.nahhs, sum.sugg.map(x => x.name).join(', ')]]);
  if (!String(sheet.getRange(row, CFG.FILE_ID_COL).getValue()).trim()) sheet.getRange(row, CFG.FILE_ID_COL).setValue(fileId);
  return true;
}

// Finds a camera's "<Site> Tagged Photos / <camera>" folder WITHOUT creating anything.
function findTaggedFolderId_(cam) {
  const key = 'f_' + hash_(cam.uid);
  let rec = null;
  try { rec = JSON.parse(props_().getProperty(key) || 'null'); } catch (e) {}
  if (rec && rec.tagged && folderOk_(rec.tagged)) return rec.tagged;
  const ov = folderOverrides_(cam);
  if (ov.taggedCam) return folderOk_(ov.taggedCam) ? ov.taggedCam : '';
  let parent = null;
  if (ov.taggedSite) {
    if (folderOk_(ov.taggedSite)) parent = DriveApp.getFolderById(ov.taggedSite);
  } else {
    const semFolder = getSemesterFolder_(cam.semester);
    const target = norm_(cam.site);
    parent = listChildFolders_(semFolder).find(x => norm_(x.getName()).indexOf(target) >= 0 && /tagged/i.test(x.getName())) || null;
  }
  if (!parent) return '';
  const it = parent.getFoldersByName(cam.uid.replace(/[\/\\]/g, '_') + '_' + cam.site + '_Tagged Photos');
  if (!it.hasNext()) return '';
  const id = it.next().getId();
  if (rec && rec.media) { rec.tagged = id; props_().setProperty(key, JSON.stringify(rec)); }
  return id;
}

// Cameras this signed-in person may browse: completed ones (instructors also see in-progress ones).
function viewerCams_(S) {
  const keys = activeSemesterKeys_();
  return readDataRows_().rows.filter(c => semesterMatches_(c.semester, keys) && (c.complete || S.admin));
}

function viewerCam_(S, uid) {
  const cam = viewerCams_(S).find(c => c.uid === String(uid));
  if (!cam) throw new Error('That camera is not available to view.');
  return cam;
}

function viewerLogin(nameStr, pinStr) {
  const realName = checkCredentials_(nameStr, pinStr);
  const admin = isAdmin_(realName);
  const sess = newSession_({ viewer: true, name: realName, admin: admin, t: Date.now() });
  return { sess: sess, token: ScriptApp.getOAuthToken(), name: realName, admin: admin, species: getSpecies_() };
}

// Cameras plus the Master Data info (species, time) of every tagged photo, in one pass.
function viewerIndex(sess) {
  const S = requireViewer_(sess);
  const cams = viewerCams_(S);
  const byUid = {};
  const cameras = cams.map(c => {
    byUid[c.uid] = [];
    return { key: c.uid, label: 'Cam ' + c.displayCam + ' (' + c.site + ')' +
      ((c.deployment && c.collection) ? ' ' + c.deployment + ' to ' + c.collection : ''),
      site: c.site, complete: c.complete, reviewer: c.reviewer };
  });
  const sheet = masterSheet_(false);
  if (sheet && sheet.getLastRow() >= 2) {
    const width = Math.min(CFG.FILE_ID_COL, sheet.getMaxColumns());
    const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
    vals.forEach(r => {
      const list = byUid[String(r[4])];
      if (!list) return;
      list.push({
        id: width >= CFG.FILE_ID_COL ? String(r[CFG.FILE_ID_COL - 1] || '') : '',
        name: String(r[6] || ''),
        species: String(r[7] || '').split(',').map(x => x.trim()).filter(Boolean),
        time: (width > 10 ? stampText_(r[10]) : '') || (width > 9 ? stampText_(r[9]) : ''),
      });
    });
  }
  return { cameras: cameras, index: byUid, species: getSpecies_() };
}

// Photos in one camera's tagged folder, with thumbnails and the star / Nahhh / suggestion totals.
function viewerPhotos(sess, uid) {
  const S = requireViewer_(sess);
  const cam = viewerCam_(S, uid);
  const taggedId = findTaggedFolderId_(cam);
  if (!taggedId) return [];
  const fb = readFeedbackAll_();
  const photos = [];
  let pageToken = null;
  do {
    const params = {
      q: "'" + taggedId + "' in parents and trashed = false",
      pageSize: 500,
      fields: 'nextPageToken, files(id, name, mimeType, thumbnailLink)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    };
    if (pageToken) params.pageToken = pageToken;
    const res = Drive.Files.list(params);
    (res.files || []).forEach(f => {
      if (f.mimeType.indexOf('image/') !== 0 && f.mimeType.indexOf('video/') !== 0) return;
      photos.push({
        id: f.id, name: f.name,
        thumb: f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s300') : '',
        big: f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s1200') : '',
        isVideo: f.mimeType.indexOf('video/') === 0,
        fb: summarizeFeedback_(fb.byId[f.id], S.name, S.admin),
      });
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  return photos;
}

// Saves this student's star / "Nahhh" / suggested IDs for one photo and refreshes the Master Data columns.
// fb = { star: bool, nahh: bool, suggested: [species, ...] }   (suggestions only count while nahh is on)
function setPhotoFeedback(sess, uid, fileId, fb) {
  const S = requireViewer_(sess);
  fileId = String(fileId || '');
  try {
    const cam = viewerCam_(S, uid);
    const taggedId = findTaggedFolderId_(cam);
    let meta;
    try { meta = Drive.Files.get(fileId, { fields: 'id,name,parents,trashed', supportsAllDrives: true }); }
    catch (e) { return { error: 'That photo could not be found.' }; }
    if (meta.trashed || !taggedId || (meta.parents || [])[0] !== taggedId) {
      return { error: 'That photo is not in this camera\'s tagged folder.' };
    }
    const star = !!(fb && fb.star);
    const nahh = !!(fb && fb.nahh);
    let sugg = [];
    if (nahh) {
      const seen = {};
      sugg = (Array.isArray(fb && fb.suggested) ? fb.suggested : [])
        .map(x => String(x).trim().replace(/\s+/g, ' ')).filter(Boolean)
        .filter(x => { const k = x.toLowerCase(); if (seen[k]) return false; seen[k] = true; return true; })
        .slice(0, CFG.MAX_SUGGESTIONS);
      if (!sugg.every(validSpecies_)) return { error: 'One of the suggested species names is not valid.' };
      const canon = canonicalizeSpecies_(sugg);            // before the lock: it takes the lock itself
      sugg = sugg.map(x => canon.get(x.toLowerCase()) || x);
    }

    let sum;
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = feedbackSheet_(true);
      const last = sh.getLastRow();
      const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 7).getValues() : [];
      const meKey = S.name.toLowerCase();
      let at = -1;
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0]) === fileId && String(vals[i][2]).toLowerCase() === meKey) { at = i; break; }
      }
      const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      const empty = !star && !nahh;
      const rowVals = [fileId, cam.uid, S.name, star ? 1 : 0, nahh ? 1 : 0, sugg.join(' | '), stamp];
      if (at >= 0) {
        if (empty) { sh.deleteRow(at + 2); vals.splice(at, 1); }
        else { sh.getRange(at + 2, 1, 1, 7).setValues([rowVals]); vals[at] = rowVals; }
      } else if (!empty) {
        sh.getRange(last + 1, 1, 1, 7).setNumberFormat('@');
        sh.getRange(last + 1, 1, 1, 7).setValues([rowVals]);
        vals.push(rowVals);
      }
      const rows = vals.filter(v => String(v[0]) === fileId).map(v => ({
        student: String(v[2]), star: Number(v[3]) === 1, nahh: Number(v[4]) === 1, sugg: splitSugg_(v[5]),
      }));
      sum = summarizeFeedback_(rows, S.name, S.admin);
      try { writeMasterFeedback_(fileId, meta.name, sum); } catch (e) {}
    } finally {
      lock.releaseLock();
    }
    return { ok: true, fb: sum, speciesList: getSpecies_() };
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  }
}

// Removes the Nahhh flags and suggestions on a photo (stars stay). Used after its ID is changed.
function clearDoubt_(fileId, fileName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = feedbackSheet_(false);
    let rows = [];
    if (sh && sh.getLastRow() >= 2) {
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {            // bottom up so row numbers stay valid
        if (String(vals[i][0]) !== fileId) continue;
        const v = vals[i];
        if (Number(v[3]) === 1) {
          sh.getRange(i + 2, 5, 1, 2).setValues([[0, '']]);
          rows.push({ student: String(v[2]), star: true, nahh: false, sugg: [] });
        } else {
          sh.deleteRow(i + 2);
        }
      }
    }
    writeMasterFeedback_(fileId, fileName, summarizeFeedback_(rows, ''));
  } finally {
    lock.releaseLock();
  }
}

// Deletes every feedback row of a photo (its tag no longer exists).
function deleteFeedbackFor_(fileId) {
  const sh = feedbackSheet_(false);
  if (!sh || sh.getLastRow() < 2) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]) === fileId) sh.deleteRow(i + 2);
  } finally {
    lock.releaseLock();
  }
}

// Instructor only: officially change a photo's species from the "View tagged photos" window.
function adminChangeTag(sess, uid, fileId, species) {
  const S = requireViewer_(sess);
  if (!S.admin) return { error: 'Only the instructor can change a tag from this window.' };
  let cam, taggedId;
  try {
    cam = viewerCam_(S, uid);
    taggedId = findTaggedFolderId_(cam);
  } catch (e) { return { error: String(e && e.message ? e.message : e) }; }
  const r = changeTagCore_({ cam: cam, taggedId: taggedId, name: cam.reviewer || S.name }, fileId, species);
  if (r.error) return r;
  const fb = readFeedbackAll_();
  r.fb = summarizeFeedback_(fb.byId[String(fileId)], S.name, true);
  return r;
}

/* ------------------------------------------------------------------ */
/* Instructor tools: camera setup, folder choices, review any camera   */
/* (all of these require an instructor sign-in; nothing here runs for  */
/* students, and nothing runs on its own)                              */
/* ------------------------------------------------------------------ */

const SETUP_COLS = { site: 1, semester: 2, camId: 3, deployment: 4, collection: 5, lat: 6, lon: 7, brand: 8, reviewer: 9, complete: 11 };
const SETUP_IDENTITY = ['site', 'semester', 'camId', 'deployment', 'collection'];
const SETUP_LABELS = {
  site: 'Site', semester: 'Semester', camId: 'Camera', deployment: 'Deployment date', collection: 'Collection date',
  lat: 'GPS N', lon: 'GPS W', brand: 'Brand', reviewer: 'Reviewer', complete: 'Complete',
};

function requireAdminViewer_(sess) {
  const S = requireViewer_(sess);
  if (!S.admin) throw new Error('Instructor access only.');
  return S;
}

function stampNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }

// Only ever appends. The tab is created on first use.
function appendSheetRows_(name, headers, rows) {
  if (!rows.length) return;
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  const start = sh.getLastRow() + 1;
  const need = start + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  const rg = sh.getRange(start, 1, rows.length, headers.length);
  rg.setNumberFormat('@');
  rg.setValues(rows);
}

function foldersSheet_(createIfMissing) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CFG.FOLDERS_SHEET);
  if (!sh && createIfMissing) {
    sh = ss.insertSheet(CFG.FOLDERS_SHEET);
    sh.getRange(1, 1, sh.getMaxRows(), CFG.FOLDERS_HEADERS.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, CFG.FOLDERS_HEADERS.length).setValues([CFG.FOLDERS_HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Folders chosen in Camera setup. Read-only; returns empty settings when the tab does not exist.
function readFolderSettings_() {
  const out = { site: {}, camera: {} };
  const sh = foldersSheet_(false);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach((r, i) => {
    const scope = String(r[0]).trim().toLowerCase();
    const key = String(r[1]).trim();
    if (!key) return;
    const rec = {
      row: i + 2,
      photosId: String(r[2] || '').trim(), photosName: String(r[3] || ''),
      taggedId: String(r[4] || '').trim(), taggedName: String(r[5] || ''),
    };
    if (scope === 'site') out.site[norm_(key)] = rec;
    else if (scope === 'camera') out.camera[key] = rec;
  });
  return out;
}

// What Camera setup says about this camera's folders. Camera choices win over site choices; anything
// left blank falls back to the automatic search.
function folderOverrides_(cam) {
  const fs = readFolderSettings_();
  const c = fs.camera[cam.uid] || {};
  const st = fs.site[norm_(cam.site)] || {};
  return { photosCam: c.photosId || '', taggedCam: c.taggedId || '', photosSite: st.photosId || '', taggedSite: st.taggedId || '' };
}

function masterCounts_() {
  const out = {};
  const sheet = masterSheet_(false);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getMaxColumns() < 5) return out;
  sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues().forEach(r => {
    const u = String(r[0] || '');
    if (u) out[u] = (out[u] || 0) + 1;
  });
  return out;
}

function rowUid_(r) {
  return String(r[11] || '').trim() || (String(r[2] || '').trim() + '_' + String(r[3] || '').trim() + '_' + String(r[4] || '').trim());
}

// Everything the "Camera setup" and "Review cameras" tabs show. Read-only.
function adminGetSetup(sess) {
  const S = requireAdminViewer_(sess);
  const data = readDataRows_();
  const counts = masterCounts_();
  const fs = readFolderSettings_();
  const keys = activeSemesterKeys_();
  const active = String(getSpreadsheet_().getSheetByName('Student Names').getRange('A1').getValue()).trim();
  const sites = {}, semesters = [];
  const cameras = data.rows.map(c => {
    if (semesters.indexOf(c.semester) < 0) semesters.push(c.semester);
    const sk = norm_(c.site);
    if (!sites[sk]) {
      const st = fs.site[sk] || {};
      sites[sk] = {
        site: c.site,
        photos: st.photosId ? { id: st.photosId, name: st.photosName } : null,
        tagged: st.taggedId ? { id: st.taggedId, name: st.taggedName } : null,
      };
    }
    const cf = fs.camera[c.uid] || {};
    const tagged = counts[c.uid] || 0;
    const why = [];
    if (c.reviewer) why.push('has a reviewer');
    if (c.complete) why.push('is complete');
    if (tagged) why.push('has ' + tagged + ' tagged photo' + (tagged === 1 ? '' : 's'));
    return {
      uid: c.uid, row: c.row, site: c.site, semester: c.semester, camId: c.camId,
      deployment: c.deployment, collection: c.collection, lat: c.lat, lon: c.lon, brand: c.brand,
      reviewer: c.reviewer, complete: c.complete,
      active: semesterMatches_(c.semester, keys),
      tagged: tagged, locked: why.length > 0, lockReason: why.join(', '),
      photos: cf.photosId ? { id: cf.photosId, name: cf.photosName } : null,
      taggedFolder: cf.taggedId ? { id: cf.taggedId, name: cf.taggedName } : null,
    };
  });
  return {
    name: S.name, active: active, cameras: cameras, semesters: semesters,
    sites: Object.keys(sites).map(k => sites[k]),
  };
}

function cleanSetupField_(field, raw) {
  const v = String(raw == null ? '' : raw).trim();
  const label = SETUP_LABELS[field];
  const text = (max, required) => {
    if (!v) return required ? { error: label + ' is required.' } : { value: '', display: '' };
    if (v.length > max) return { error: label + ' is too long.' };
    if (/^[=+\-@]/.test(v) || /[\u0000-\u001f]/.test(v)) return { error: label + ' contains characters that are not allowed.' };
    return { value: v, display: v };
  };
  switch (field) {
    case 'site': case 'semester': return text(60, true);
    case 'brand': return text(60, false);
    case 'camId': {
      const t = text(40, true);
      if (t.error) return t;
      const x = /^\d+$/.test(v) ? 'Cam ' + v : v;
      return { value: x, display: x };
    }
    case 'deployment': case 'collection': {
      const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(v);
      if (!m) return { error: label + ' must be a date like 10/2/2026.' };
      const mo = +m[1], d = +m[2], y = +m[3];
      const dt = new Date(y, mo - 1, d);
      if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return { error: label + ' is not a real date.' };
      return { value: dt, display: mo + '/' + d + '/' + y, fmt: 'M/d/yyyy' };
    }
    case 'lat': case 'lon': {
      if (!/^-?\d+(\.\d+)?$/.test(v)) return { error: label + ' must be a number like 37.09181.' };
      const n = parseFloat(v);
      if (Math.abs(n) > (field === 'lat' ? 90 : 180)) return { error: label + ' is out of range.' };
      return { value: n, display: String(n) };
    }
    case 'reviewer':
      return v === '' ? { value: '', display: '' } : { error: 'Reviewer can only be cleared (released) here.' };
    case 'complete':
      return (v === 'Yes' || v === 'No') ? { value: v, display: v } : { error: 'Complete must be Yes or No.' };
  }
  return { error: 'Unknown field.' };
}

// Saves camera data changes. Rules that keep your current data safe:
//  - only the specific cells that changed are written, never a whole row or table;
//  - every change carries the value the instructor saw ("old"); if the sheet differs now, nothing is saved;
//  - camera / site / semester / dates are locked for cameras that have a reviewer, are complete, or have tagged photos;
//  - new cameras are appended below the last row (other columns, such as "How to use", are left alone);
//  - the whole batch is validated first, so it is saved completely or not at all;
//  - every change is written to the "Setup Change Log" tab with the old and new value.
function adminSaveSetup(sess, payload) {
  const S = requireAdminViewer_(sess);
  const edits = Array.isArray(payload && payload.edits) ? payload.edits : [];
  const adds = Array.isArray(payload && payload.adds) ? payload.adds : [];
  if (!edits.length && !adds.length) return { ok: true, changed: 0, added: [] };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSpreadsheet_().getSheetByName('Data');
    const vals = sheet.getDataRange().getDisplayValues();          // fresh read, under the same lock claims use
    const camByRow = {};
    for (let i = 1; i < vals.length; i++) {
      const c = rowToCam_(vals[i], i + 1);
      if (c.site && c.camId) camByRow[i + 1] = c;
    }
    const counts = masterCounts_();
    const problems = [];
    const writes = [];
    const proposed = {};                                            // row -> display values after the edits
    const touchedIdentity = {};
    const stamp = stampNow_();

    edits.forEach(ed => {
      const row = parseInt(ed && ed.row, 10);
      const cam = camByRow[row];
      if (!cam || cam.uid !== String(ed.uid)) {
        problems.push('The camera in row ' + row + ' changed or moved in the sheet. Refresh the page and try again.');
        return;
      }
      const label = cam.camId + ' (' + cam.site + ')';
      const why = [];
      if (cam.reviewer) why.push('it has a reviewer');
      if (cam.complete) why.push('it is complete');
      if (counts[cam.uid]) why.push('it has tagged photos');
      const rv = vals[row - 1];
      proposed[row] = proposed[row] || rv.slice();
      Object.keys(ed.fields || {}).forEach(field => {
        const col = SETUP_COLS[field];
        const ch = ed.fields[field] || {};
        if (!col) { problems.push(label + ': unknown field "' + field + '".'); return; }
        const current = String(rv[col - 1] == null ? '' : rv[col - 1]).trim();
        const oldSeen = String(ch.old == null ? '' : ch.old).trim();
        const sameAsSeen = field === 'complete' ? ((oldSeen.toLowerCase() === 'yes') === (current.toLowerCase() === 'yes')) : (oldSeen === current);
        if (!sameAsSeen) {
          problems.push(label + ': ' + SETUP_LABELS[field] + ' was changed in the sheet after you opened this page (it is now "' + current + '"). Refresh and try again.');
          return;
        }
        if (SETUP_IDENTITY.indexOf(field) >= 0 && why.length) {
          problems.push(label + ': ' + SETUP_LABELS[field] + ' is locked because ' + why.join(' and ') + '.');
          return;
        }
        const clean = cleanSetupField_(field, ch.new);
        if (clean.error) { problems.push(label + ': ' + clean.error); return; }
        if (field === 'complete' ? ((clean.value === 'Yes') === (current.toLowerCase() === 'yes')) : (clean.display === current)) return;
        writes.push({ row: row, col: col, value: clean.value, fmt: clean.fmt || '', field: field, cam: cam, old: current, now: clean.display });
        proposed[row][col - 1] = clean.display;
        if (SETUP_IDENTITY.indexOf(field) >= 0) touchedIdentity[row] = true;
      });
    });

    const newRows = [];
    let nextRow = sheet.getLastRow() + 1;
    adds.forEach((ad, i) => {
      const f = (ad && ad.fields) || {};
      const clean = {};
      let ok = true;
      ['site', 'semester', 'camId', 'deployment', 'collection', 'lat', 'lon', 'brand'].forEach(field => {
        const c = cleanSetupField_(field, f[field]);
        if (c.error) { problems.push('New camera ' + (i + 1) + ': ' + c.error); ok = false; } else clean[field] = c;
      });
      if (!ok) return;
      newRows.push({
        index: ad.index == null ? i : ad.index, row: nextRow++, clean: clean,
        uid: clean.camId.display + '_' + clean.deployment.display + '_' + clean.collection.display,
      });
    });

    // The same camera number and dates may only appear once (progress, tags and folders are keyed by it).
    const owners = {};
    Object.keys(camByRow).forEach(r => {
      const uid = touchedIdentity[r] ? rowUid_(proposed[r]) : camByRow[r].uid;
      (owners[uid] = owners[uid] || []).push(String(r));
    });
    newRows.forEach(n => { (owners[n.uid] = owners[n.uid] || []).push('new'); });
    Object.keys(touchedIdentity).forEach(r => {
      const uid = rowUid_(proposed[r]);
      if (owners[uid].length > 1) problems.push(camByRow[r].camId + ' (' + camByRow[r].site + '): another camera already has the same camera number and dates.');
    });
    newRows.forEach(n => {
      if (owners[n.uid].length > 1) problems.push('New camera ' + n.clean.camId.display + ' (' + n.clean.site.display + '): a camera with the same camera number and dates already exists.');
    });

    if (problems.length) return { error: problems.join('\n') };

    const logRows = [];
    writes.forEach(w => {
      const rg = sheet.getRange(w.row, w.col);
      rg.setValue(w.value);
      if (w.fmt) rg.setNumberFormat(w.fmt);
      let action = 'Edit camera';
      if (w.field === 'reviewer') action = 'Release reviewer';
      else if (w.field === 'complete') action = w.value === 'Yes' ? 'Mark complete' : 'Reopen camera';
      logRows.push([stamp, S.name, action, w.cam.uid, SETUP_LABELS[w.field] + ' (row ' + w.row + ')', w.old, w.now]);
    });
    newRows.forEach(n => {
      ['site', 'semester', 'camId', 'deployment', 'collection', 'lat', 'lon', 'brand'].forEach(field => {
        const c = n.clean[field];
        if (field === 'brand' && c.value === '') return;          // leave empty cells empty
        const rg = sheet.getRange(n.row, SETUP_COLS[field]);
        rg.setValue(c.value);
        if (c.fmt) rg.setNumberFormat(c.fmt);
      });
      logRows.push([stamp, S.name, 'Add camera', n.uid, 'new row ' + n.row, '', n.clean.site.display + ', ' + n.clean.semester.display + ', ' + n.clean.camId.display + ', ' + n.clean.deployment.display + ' to ' + n.clean.collection.display]);
    });
    appendSheetRows_(CFG.LOG_SHEET, CFG.LOG_HEADERS, logRows);
    return { ok: true, changed: writes.length + newRows.length, added: newRows.map(n => ({ index: n.index, uid: n.uid, row: n.row })) };
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

// Saves folders chosen in Camera setup into the "Camera Folders" tab (a tab this app owns; created on first save).
// changes: [{ scope: 'site'|'camera', key: <site name or camera id>, kind: 'photos'|'tagged', id: <folder id or '' for automatic> }]
function adminSaveFolders(sess, changes) {
  const S = requireAdminViewer_(sess);
  changes = Array.isArray(changes) ? changes : [];
  if (!changes.length) return { ok: true, changed: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = readDataRows_();
    const byUid = {}, siteName = {};
    data.rows.forEach(c => { byUid[c.uid] = c; siteName[norm_(c.site)] = c.site; });
    const todo = [];
    for (let i = 0; i < changes.length; i++) {
      const ch = changes[i] || {};
      const scope = ch.scope === 'camera' ? 'Camera' : (ch.scope === 'site' ? 'Site' : '');
      const kind = ch.kind === 'tagged' ? 'tagged' : (ch.kind === 'photos' ? 'photos' : '');
      const key = String(ch.key || '').trim();
      if (!scope || !kind || !key) return { error: 'Invalid folder change.' };
      if (scope === 'Camera' && !byUid[key]) return { error: 'Camera not found: ' + key };
      if (scope === 'Site' && !siteName[norm_(key)]) return { error: 'Site not found: ' + key };
      const id = String(ch.id || '').trim();
      let name = '';
      if (id) {
        if (!validDriveId_(id) || id === '@shared') return { error: 'Invalid folder.' };
        let m;
        try { m = Drive.Files.get(id, { fields: 'id,name,mimeType,trashed', supportsAllDrives: true }); }
        catch (e) { return { error: 'Could not open the chosen folder. Is it shared with the account that runs this app?' }; }
        if (m.trashed || m.mimeType !== FOLDER_MIME) return { error: 'That is not a folder.' };
        name = m.name;
      }
      todo.push({ scope: scope, kind: kind, key: scope === 'Site' ? siteName[norm_(key)] : key, id: id, name: name });
    }
    const sh = foldersSheet_(true);
    const last = sh.getLastRow();
    const existing = last >= 2 ? sh.getRange(2, 1, last - 1, 6).getValues() : [];
    const rowOf = {};
    existing.forEach((r, i) => { rowOf[String(r[0]).trim().toLowerCase() + '|' + (String(r[0]).trim().toLowerCase() === 'site' ? norm_(r[1]) : String(r[1]).trim())] = i + 2; });
    const cur = {};                                                  // row -> [photosId, photosName, taggedId, taggedName]
    existing.forEach((r, i) => { cur[i + 2] = [String(r[2] || ''), String(r[3] || ''), String(r[4] || ''), String(r[5] || '')]; });
    let nextRow = last + 1;
    const plan = [];
    for (let i = 0; i < todo.length; i++) {
      const t = todo[i];
      const rk = t.scope.toLowerCase() + '|' + (t.scope === 'Site' ? norm_(t.key) : t.key);
      let row = rowOf[rk];
      if (!row) { row = nextRow++; rowOf[rk] = row; cur[row] = ['', '', '', '']; plan.push({ row: row, fresh: true, scope: t.scope, key: t.key }); }
      const c = cur[row];
      const before = t.kind === 'photos' ? c[1] : c[3];
      if (t.kind === 'photos') { c[0] = t.id; c[1] = t.name; } else { c[2] = t.id; c[3] = t.name; }
      plan.push({ row: row, t: t, before: before });
    }
    for (const row in cur) {
      if (cur[row][0] && cur[row][0] === cur[row][2]) return { error: 'The photo folder and the tagged-photos folder cannot be the same folder.' };
    }
    const stamp = stampNow_();
    const logRows = [];
    const touched = {};
    plan.forEach(p => {
      if (p.fresh) {
        sh.getRange(p.row, 1, 1, 2).setNumberFormat('@');
        sh.getRange(p.row, 1, 1, 2).setValues([[p.scope, p.key]]);
        return;
      }
      touched[p.row] = true;
      const t = p.t;
      logRows.push([stamp, S.name, 'Folder choice', t.scope === 'Site' ? 'Site: ' + t.key : t.key, (t.kind === 'photos' ? 'Photos folder' : 'Tagged folder'), p.before || '(automatic)', t.name || '(automatic)']);
    });
    Object.keys(touched).forEach(row => {
      const c = cur[row];
      sh.getRange(+row, 3, 1, 6).setNumberFormat('@');
      sh.getRange(+row, 3, 1, 6).setValues([[c[0], c[1], c[2], c[3], stamp, S.name]]);
    });
    appendSheetRows_(CFG.LOG_SHEET, CFG.LOG_HEADERS, logRows);
    // Forget remembered folders so the next login uses the new choice.
    todo.forEach(t => {
      if (t.scope === 'Camera') props_().deleteProperty('f_' + hash_(t.key));
      else data.rows.forEach(c => { if (norm_(c.site) === norm_(t.key)) props_().deleteProperty('f_' + hash_(c.uid)); });
    });
    return { ok: true, changed: todo.length };
  } catch (e) {
    return { error: String(e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

function validDriveId_(id) { return id === 'root' || id === '@shared' || /^[A-Za-z0-9_-]{5,}$/.test(String(id)); }
function qEsc_(v) { return String(v).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

// ---- read-only folder browser used by the folder picker ----
function adminBrowse(sess, folderId) {
  requireAdminViewer_(sess);
  folderId = String(folderId || '');
  const out = { id: folderId, name: '', crumbs: [], folders: [] };
  if (!folderId) {
    out.name = 'Locations';
    out.folders = [{ id: 'root', name: 'My Drive' }, { id: '@shared', name: 'Shared with me' }];
    try {
      const d = Drive.Drives.list({ pageSize: 100, fields: 'drives(id,name)' });
      (d.drives || []).forEach(x => out.folders.push({ id: x.id, name: x.name + ' (shared drive)' }));
    } catch (e) {}
    return out;
  }
  if (!validDriveId_(folderId)) return { error: 'Invalid folder.' };
  let q;
  if (folderId === '@shared') {
    out.name = 'Shared with me';
    q = "sharedWithMe = true and mimeType = '" + FOLDER_MIME + "' and trashed = false";
  } else {
    let meta;
    try { meta = Drive.Files.get(folderId, { fields: 'id,name,parents,mimeType,trashed', supportsAllDrives: true }); }
    catch (e) { return { error: 'That folder could not be opened.' }; }
    if (meta.trashed || meta.mimeType !== FOLDER_MIME) return { error: 'That is not a folder.' };
    out.id = meta.id;
    out.name = folderId === 'root' ? 'My Drive' : meta.name;
    let cur = meta, guard = 0;
    while (cur && cur.parents && cur.parents.length && guard++ < 12) {
      try {
        const p = Drive.Files.get(cur.parents[0], { fields: 'id,name,parents', supportsAllDrives: true });
        out.crumbs.unshift({ id: p.id, name: p.name });
        cur = p;
      } catch (e) { break; }
    }
    q = "'" + qEsc_(meta.id) + "' in parents and mimeType = '" + FOLDER_MIME + "' and trashed = false";
  }
  const res = Drive.Files.list({
    q: q, pageSize: 1000, orderBy: 'name', fields: 'files(id, name)',
    supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
  });
  out.folders = (res.files || []).map(f => ({ id: f.id, name: f.name }));
  if (folderId !== '@shared') {
    const files = Drive.Files.list({
      q: "'" + qEsc_(out.id) + "' in parents and mimeType != '" + FOLDER_MIME + "' and trashed = false", pageSize: 1000,
      fields: 'nextPageToken, files(id, name)', supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
    });
    out.media = (files.files || []).filter(f => CFG.MEDIA_EXT.test(f.name)).length;
    out.mediaMore = !!files.nextPageToken;
  }
  return out;
}

function adminSearchFolders(sess, text) {
  requireAdminViewer_(sess);
  const t = String(text || '').trim();
  if (t.length < 2) return [];
  const res = Drive.Files.list({
    q: "mimeType = '" + FOLDER_MIME + "' and trashed = false and name contains '" + qEsc_(t) + "'",
    pageSize: 25, fields: 'files(id, name, parents)',
    supportsAllDrives: true, includeItemsFromAllDrives: true, corpora: 'allDrives',
  });
  const names = {};
  return (res.files || []).map(f => {
    const pid = (f.parents || [])[0];
    let parent = '';
    if (pid) {
      if (!(pid in names)) {
        try { names[pid] = Drive.Files.get(pid, { fields: 'name', supportsAllDrives: true }).name || ''; } catch (e) { names[pid] = ''; }
      }
      parent = names[pid];
    }
    return { id: f.id, name: f.name, parent: parent };
  });
}

// Counts the photos and videos under a folder (stops counting at `cap`). Read-only.
function countMedia_(rootId, cap, maxCalls) {
  cap = cap || 5000;
  maxCalls = maxCalls || 40;
  const queue = [rootId];
  let n = 0, subs = 0, capped = false, calls = 0;
  while (queue.length && !capped) {
    if (calls >= maxCalls) { capped = true; break; }
    const cur = queue.shift();
    let pageToken = null;
    do {
      const params = {
        q: "'" + qEsc_(cur) + "' in parents and trashed = false", pageSize: 1000,
        fields: 'nextPageToken, files(id, name, mimeType)', supportsAllDrives: true, includeItemsFromAllDrives: true,
      };
      if (pageToken) params.pageToken = pageToken;
      calls++;
      const res = Drive.Files.list(params);
      (res.files || []).forEach(f => {
        if (f.mimeType === FOLDER_MIME) { if (!/tagged/i.test(f.name)) { queue.push(f.id); subs++; } }
        else if (CFG.MEDIA_EXT.test(f.name)) n++;
      });
      pageToken = res.nextPageToken;
      if (n >= cap) { capped = true; break; }
    } while (pageToken);
  }
  return { count: n, capped: capped, subfolders: subs };
}

// Finds where a camera's folders are WITHOUT creating or changing anything.
function peekFolders_(cam) {
  const out = { source: 'auto', media: null, tagged: null, taggedSource: 'auto', error: '' };
  const ov = folderOverrides_(cam);
  try {
    let mf = null;
    if (ov.photosCam) {
      out.source = 'camera';
      if (!folderOk_(ov.photosCam)) throw new Error('The chosen photo folder is missing or in the trash.');
      mf = DriveApp.getFolderById(ov.photosCam);
    } else if (ov.photosSite) {
      out.source = 'site';
      if (!folderOk_(ov.photosSite)) throw new Error('The chosen site photo folder is missing or in the trash.');
      mf = pickMediaFolder_(DriveApp.getFolderById(ov.photosSite), cam);
      if (!mf) throw new Error('No folder for ' + cam.camId + ' inside the chosen site photo folder.');
    } else {
      const siteFolder = getSiteFolder_(cam.semester, cam.site);
      const reviewBase = getTargetSubfolder_(siteFolder, 'cameraphotostoreview');
      if (!reviewBase) throw new Error('Could not find CAMERA PHOTOS TO REVIEW folder in ' + siteFolder.getName());
      mf = pickMediaFolder_(reviewBase, cam);
      if (!mf) throw new Error('Could not locate media folder for Camera ' + cam.displayCam);
    }
    out.media = { id: mf.getId(), name: mf.getName() };
  } catch (e) {
    out.error = String(e && e.message ? e.message : e);
    return out;
  }
  try {
    const tid = findTaggedFolderId_(cam);
    if (tid) out.tagged = { id: tid, name: Drive.Files.get(tid, { fields: 'id,name', supportsAllDrives: true }).name };
    out.taggedSource = ov.taggedCam ? 'camera' : (ov.taggedSite ? 'site' : 'auto');
  } catch (e) {}
  return out;
}

function adminCheckCamera(sess, uid) {
  requireAdminViewer_(sess);
  const cam = readDataRows_().rows.find(c => c.uid === String(uid));
  if (!cam) return { error: 'Camera not found. Refresh the page.' };
  const p = peekFolders_(cam);
  if (p.media) { const c = countMedia_(p.media.id, 5000); p.count = c.count; p.capped = c.capped; }
  return p;
}

// ---- review any camera as the instructor ----
// Opens a camera exactly like a student's review, except nothing is claimed or written: the Reviewer column
// and the sign-up sheet are not touched, and no folder is created until the first photo is tagged.
function instructorPayload_(name, cam, withFiles) {
  const ids = getFolderIds_(cam, false);
  let taggedId = '';
  try { taggedId = findTaggedFolderId_(cam); } catch (e) { taggedId = ''; }
  const tree = walkMedia_(ids.media, ids.mediaName, taggedId, withFiles);
  const sess = newSession_({
    name: name, instr: true, mediaId: ids.media, taggedId: taggedId, folders: tree.folders,
    cam: {
      uid: cam.uid, site: cam.site, semester: cam.semester, camId: cam.camId, displayCam: cam.displayCam,
      lat: cam.lat, lon: cam.lon, deployment: cam.deployment, collection: cam.collection,
    },
    t: Date.now(),
  });
  return { sess: sess, token: ScriptApp.getOAuthToken(), tree: tree };
}

function instructorOpenCamera(sess, uid) {
  const A = requireAdminViewer_(sess);
  const cam = readDataRows_().rows.find(c => c.uid === String(uid));
  if (!cam) throw new Error('Camera not found. Refresh the page and try again.');
  const p = instructorPayload_(A.name, cam, true);
  return {
    sess: p.sess, token: p.token, config: publicConfig_(cam, A.name), files: p.tree.files,
    startIndex: startIndexFor_(cam, A.name, p.tree.files), skipped: p.tree.skipped,
  };
}

// Used by the browser to get a fresh instructor session after the old one expired.
function instructorRelogin(nameStr, pinStr, uid) {
  const realName = checkCredentials_(nameStr, pinStr);
  if (!isAdmin_(realName)) throw new Error('Instructor access only.');
  const cam = readDataRows_().rows.find(c => c.uid === String(uid));
  if (!cam) throw new Error('Camera not found.');
  const p = instructorPayload_(realName, cam, false);
  return { sess: p.sess, token: p.token };
}

// Instructor sessions can open a camera that has no tagged folder yet; it is created now, on the first tag.
function ensureTaggedFolder_(sess, S) {
  const rec = getFolderIds_(S.cam, true);
  S.taggedId = rec.tagged;
  cache_().put('sess_' + sess, JSON.stringify(S), CFG.SESSION_TTL);
}

/* ------------------------------------------------------------------ */
/* Instructor helper - run manually from the Apps Script editor        */
/* ------------------------------------------------------------------ */

/**
 * Run this from the Apps Script editor (select diagnoseSite, click Run, then View > Logs / Execution log)
 * to see, for every camera of a site, which Drive folder the app picked and how many files it can see.
 * Change the default below, or call diagnoseSite('YOUR_SITE') from another function.
 */
function diagnoseSite(site) {
  site = site || 'YOUR_SITE';
  const rows = readDataRows_().rows.filter(c => norm_(c.site) === norm_(site));
  Logger.log('Sheet rows for ' + site + ': ' + rows.length);
  rows.forEach(cam => {
    try {
      const siteFolder = getSiteFolder_(cam.semester, cam.site);
      const ids = getFolderIds_(cam, false);
      const t = walkMedia_(ids.media, ids.mediaName, ids.tagged, true);
      const withThumb = t.files.filter(f => f.thumbUrl).length;
      Logger.log(cam.uid + ' | site folder: "' + siteFolder.getName() + '" | media folder: "' + ids.mediaName +
        '" | usable files: ' + t.files.length + ' | with Drive thumbnail: ' + withThumb +
        ' | unsupported files skipped: ' + t.skipped + ' | subfolders: ' + (Object.keys(t.folders).length - 1));
    } catch (e) {
      Logger.log(cam.uid + ' | ERROR: ' + e.message);
    }
  });
}

/**
 * Optional: run once from the editor to add the Stars / ID_Doubt / Suggested_ID columns to Master Data
 * and create the "Photo Feedback" tab right away. (They are also created automatically on first use.)
 */
function setupFeedbackColumns() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureMasterColumns_(masterSheet_(true));
    feedbackSheet_(true);
  } finally {
    lock.releaseLock();
  }
  Logger.log('Master Data columns and the Photo Feedback tab are ready.');
}

// Run this after you move, rename, or re-create camera folders or the species sheet.
function clearCaches() {
  const p = props_();
  const all = p.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('f_') === 0 || k === 'SPECIES_FILE_ID') p.deleteProperty(k);
  });
  cache_().remove('species');
  return 'Cleared remembered folder IDs and the species cache.';
}

