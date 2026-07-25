/* ==================================================
1. MAIN PORTAL CONTROLLER
================================================== */




var TOOLS_BUILD_VERSION = 'lesson-plan-v2';











/* ==================================================
FIXED: DROPDOWN DATA (PURE JS - NO AUTH CRASH)
================================================== */






/* ==================================================
3. CALENDAR DATA SOURCE (GROUPED)
================================================== */

/* ==================================================
FIXED: CALENDAR EVENTS (Uses Column G for Route)
================================================== */






/* ==================================================
   FLIGHT TIME BLOCKS
================================================== */



/* ==================================================
4. DISPATCH SAVING & FATIGUE
================================================== */











/* ==================================================
5. SUPERVISOR DASHBOARD (NO AUTH CHECK)
================================================== */


/* ==================================================
FIXED: MISSION DETAILS (PURE JS - NO AUTH CRASH)
================================================== */








// 6. HELPER ACTIONS


































/* ==================================================
FIXED: TIMELINE (PURE JS - NO AUTH CRASH)
================================================== */
















































var DUTY_APP_BUILD_VERSION = '2026.07.16.1';






/* ==================================================
FIXED: SCHEDULED MISSIONS LIST (PURE JS)
================================================== */




/* ==================================================
MISSION-WIDE DELETION & FUEL REVERSAL
================================================== */



































































/**
* Finds a specific mission by its ID to identify which aircraft is assigned.
* Used by the Dashboard to trigger the Tech Status/Squawk lookup.
*/













// Backwards-compatible wrapper: client calls saveMissionToLog








// ── Flight Following ──────────────────────────────────────────────────────────
var FF_RECIPIENTS_ = ['acompanhamento@asasdesocorro.org.br', 'supervisor.voo@asasdesocorro.org.br'];
var FF_MESSAGES_SPREADSHEET_ID_PROP_ = 'FF_MESSAGES_SPREADSHEET_ID';
var FF_MESSAGES_SPREADSHEET_ID_DEFAULT_ = '1LcpCPau14b_ugcWufBRPwrIV5q-juWB_P6XETDaP3JM';
var FF_MESSAGES_HEADERS_ = [
  'TIMESTAMP',
  'REG',
  'REG_NORM',
  'DIRECTION',
  'FROM_ADDR',
  'TO_ADDR',
  'SUBJECT',
  'BODY',
  'VERIFIED',
  'THREAD_ID',
  'MESSAGE_ID',
  'SOURCE',
  'SYNCED_AT'
];
var FF_ARCHIVE_HEADERS_ = [
  'ARCHIVED_AT',
  'ARCHIVE_ID',
  'FLIGHT_ID',
  'MISSION_ID',
  'FLIGHT_LEG_ID',
  'REG',
  'ARCHIVED_BY',
  'REASON',
  'SNAPSHOT_JSON',
  'SNAPSHOT_FILE_ID',
  'SNAPSHOT_FILE_URL',
  'SNAPSHOT_SIZE'
];








































































































  /* ==================================================
  SETUP: Create RUNWAY_WALKTHROUGH_LOG sheet if needed
  =================================================== */


  /* ================================================== 
  RUNWAY WALKTHROUGH & FEATURE EDITS
  ================================================== */



/* ==================================================
   RUNWAY DATABASE — save full entry from Pilot App
   Public (no underscore) so google.script.run can reach it.
   ================================================== */
































































































































/**
 * Apply a FIFO student debit for a completed training mission.
 * Idempotent: skips if a FLIGHT_DEBIT_FIFO entry for missionId already exists.
 * Returns { success, skipped?, reason?, totalDebited?, newBalance?, error? }
 */





// ---- Student Deposit Void / Edit ----





































































































// ─────────────────────────────────────────────────────────────────
// RUNWAY BRIEFING CARD  –  save / load
// ─────────────────────────────────────────────────────────────────

/**
 * Saves a briefingCard blob into the KNOWN_FEATURES JSON for the
 * matching runway row(s) in DB_Airports.
 * Both runway directions (e.g. RWY 10 and RWY 28) get the same card
 * so either direction can open it.
 */

/**
/**
 * Uploads an airstrip photo (base64-encoded) to Google Drive under
 * FlightApp_Airstrip_Photos/<ICAO>/ and returns a public view URL.
 */

/**
 * Returns the last `limit` takeoff departures from `icao` (matching
 * rwyIdent where available) from LOG_Flights, with performance details.
 */







// ─── Training-to-Scheduler Qualification Sync ────────────────────────────────









// ─── Training Module Course Generator ────────────────────────────────────────


// ============================================================
// LIVE POSITION TRACKING
// ============================================================
