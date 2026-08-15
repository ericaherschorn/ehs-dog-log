// ============================================================
// Tests.gs — Backend test runner for EHS Dog Log
//
// HOW TO RUN:
//   1. Open script.google.com → open this project
//   2. Select "runTests" in the function dropdown
//   3. Click ▶ Run
//   4. Open "Execution log" to see ✓ / ✗ results
//   5. After testing, run cleanupTestData() to remove __TEST_ rows
// ============================================================

function runTests() {
  var ok = 0, fail = 0;

  function assert(name, condition, detail) {
    if (condition) {
      Logger.log("✓ " + name);
      ok++;
    } else {
      Logger.log("✗ FAIL: " + name + (detail ? " — " + detail : ""));
      fail++;
    }
  }

  Logger.log("══════════════════════════════════════════");
  Logger.log("EHS Dog Log — Backend Tests");
  Logger.log("══════════════════════════════════════════");

  // ── Auth ──────────────────────────────────────────────────
  Logger.log("\n── Auth ──");
  var role = getUserRole_();
  assert("getUserRole_ returns valid role",
    ["Admin", "Regular", "Alert Access", "View", "none"].indexOf(role) !== -1,
    "got: " + role);

  // ── Profile save / update ─────────────────────────────────
  Logger.log("\n── Profile save/update ──");
  var testName = "__TEST_" + new Date().getTime();
  var createResult = saveDogProfile({
    originalDogName: testName,
    dogName:         testName,
    status:          "Active",
    breed:           "Test Breed"
  });
  assert("saveDogProfile creates dog",  createResult.action === "created",  "got: " + createResult.action);

  var updateResult = saveDogProfile({
    originalDogName: testName,
    dogName:         testName,
    status:          "Active",
    breed:           "Updated Breed"
  });
  assert("saveDogProfile updates dog",  updateResult.action === "updated",  "got: " + updateResult.action);

  // ── Photo URL preserved on save ───────────────────────────
  Logger.log("\n── Photo URL preservation ──");
  var sheet = getSpreadsheet_().getSheetByName(PROFILES_TAB);
  var data  = sheet.getDataRange().getValues();
  var testRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][PC.DOG_NAME] === testName) { testRow = i; break; }
  }
  if (testRow !== -1) {
    sheet.getRange(testRow + 1, PC.PHOTO_URL + 1).setValue("https://example.com/photo.jpg");
    saveDogProfile({ originalDogName: testName, dogName: testName, status: "Active" });
    var reloaded = sheet.getDataRange().getValues();
    assert("saveDogProfile preserves photoUrl on update",
      reloaded[testRow][PC.PHOTO_URL] === "https://example.com/photo.jpg");
  }

  // ── Log save / retrieve ───────────────────────────────────
  Logger.log("\n── Log save/retrieve ──");
  var tz    = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var logResult = saveLogEntry(today, "AM", testName, "2", "", "Fully Eaten", "30 min", "Test comment");
  assert("saveLogEntry succeeds", logResult.success === true, JSON.stringify(logResult));
  var entries = getLogForDate(today, "AM");
  var found   = entries.filter(function(e) { return e.dogName === testName; });
  assert("saveLogEntry is retrievable",       found.length === 1);
  assert("food value stored as 'Fully Eaten'", found.length && found[0].food === "Fully Eaten");

  // ── New food options round-trip ───────────────────────────
  Logger.log("\n── Food options ──");
  ["Fully Eaten", "Half Eaten", "Not Eaten"].forEach(function(val) {
    saveLogEntry(today, "AM", testName, "", "", val, "", "");
    var rows = getLogForDate(today, "AM").filter(function(e) { return e.dogName === testName; });
    assert("food '" + val + "' round-trips correctly",
      rows.length > 0 && rows[0].food === val, "got: " + (rows[0] && rows[0].food));
  });

  // ── Profile lock acquire / release ────────────────────────
  Logger.log("\n── Profile locking ──");
  var lock1 = acquireProfileLock(testName);
  assert("acquireProfileLock succeeds on free dog",         lock1.success === true);
  var lock2 = acquireProfileLock(testName); // same user re-acquiring
  assert("acquireProfileLock re-acquire by same user is ok", lock2.success === true);
  var renew = renewProfileLock(testName);
  assert("renewProfileLock succeeds",                        renew.success === true);
  var rel   = releaseProfileLock(testName);
  assert("releaseProfileLock succeeds",                      rel.success === true);
  var lock3 = acquireProfileLock(testName);
  assert("acquireProfileLock works again after release",     lock3.success === true);
  releaseProfileLock(testName);

  // ── Log lock acquire / release ────────────────────────────
  Logger.log("\n── Log locking ──");
  var llock1 = acquireLogLock(today, "AM", testName);
  assert("acquireLogLock succeeds",       llock1.success === true);
  var lrenew = renewLogLock(today, "AM", testName);
  assert("renewLogLock succeeds",         lrenew.success === true);
  var lrel   = releaseLogLock(today, "AM", testName);
  assert("releaseLogLock succeeds",       lrel.success === true);

  // ── Schema: PROFILE_HEADERS length ───────────────────────
  Logger.log("\n── Schema ──");
  assert("PROFILE_HEADERS has 28 entries (including Photo URL)",
    PROFILE_HEADERS.length === 28, "got: " + PROFILE_HEADERS.length);
  assert("PROFILE_HEADERS last entry is 'Photo URL'",
    PROFILE_HEADERS[PROFILE_HEADERS.length - 1] === "Photo URL");

  // ── Summary ───────────────────────────────────────────────
  Logger.log("\n══════════════════════════════════════════");
  Logger.log(ok + " passed · " + fail + " failed" + (fail === 0 ? " ✅ All tests passed!" : " ❌ Some tests failed."));
  Logger.log("══════════════════════════════════════════");
}

function cleanupTestData() {
  // Removes all __TEST_ rows created by runTests(). Safe to run multiple times.
  var ss = getSpreadsheet_();

  var profileSheet = ss.getSheetByName(PROFILES_TAB);
  var profileData  = profileSheet.getDataRange().getValues();
  for (var i = profileData.length - 1; i >= 1; i--) {
    if (String(profileData[i][PC.DOG_NAME]).indexOf("__TEST_") === 0) {
      profileSheet.deleteRow(i + 1);
    }
  }

  var logSheet = ss.getSheetByName(LOG_TAB);
  var logData  = logSheet.getDataRange().getValues();
  for (var j = logData.length - 1; j >= 1; j--) {
    if (String(logData[j][LC.DOG_NAME]).indexOf("__TEST_") === 0) {
      logSheet.deleteRow(j + 1);
    }
  }

  // Also clean profile history entries for test dogs
  var histSheet = ss.getSheetByName(PROFILE_HISTORY_TAB);
  if (histSheet) {
    var histData = histSheet.getDataRange().getValues();
    for (var k = histData.length - 1; k >= 1; k--) {
      // History rows have [date, email, ...profile fields]; DOG_NAME is at col 2 (0-based index 2)
      if (String(histData[k][2]).indexOf("__TEST_") === 0) {
        histSheet.deleteRow(k + 1);
      }
    }
  }

  Logger.log("Test data cleaned up.");
}
