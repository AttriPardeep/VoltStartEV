const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  LevelFormat,
  TabStopType,
  TabStopPosition,
  PageNumber,
  SimpleField,
} = require("docx");

const makeHeader = () =>
  new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: "VoltStartEV  ",
            font: "Arial",
            size: 18,
            bold: true,
            color: C.teal,
          }),
          new TextRun({
            text: "End-to-End Test Cases v1.0",
            font: "Arial",
            size: 18,
            color: C.gray,
          }),
        ],
      }),
    ],
  });
  
const makeFooter = () =>
  new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: {
          top: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: C.teal,
            space: 4,
          },
        },
        spacing: { before: 60 },
        children: [
          new TextRun({
            text: "CONFIDENTIAL — VoltStartEV Internal    Page ",
            font: "Arial",
            size: 16,
            color: C.gray,
          }),

          new SimpleField("PAGE"),
        ],
      }),
    ],
  });
const fs = require('fs');

// ── Palette ──────────────────────────────────────────────────────
const C = {
  navy:    '0F172A', slate:   '1E293B', teal:    '0D9488',
  tealDim: '0E7490', green:   '10B981', amber:   'F59E0B',
  red:     'EF4444', purple:  '7C3AED', blue:    '3B82F6',
  gray:    '64748B', sub:     '94A3B8', white:   'FFFFFF',
  border:  'CBD5E1', light:   'F0FDFA',
};

// ── Helpers ───────────────────────────────────────────────────────
const b  = (c=C.border) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const ab = (c=C.border) => ({ top: b(c), bottom: b(c), left: b(c), right: b(c) });

const sp = (pt=8) => new Paragraph({
  children: [new TextRun('')], spacing: { before: pt*20, after: 0 }
});

const h1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text: t, font: 'Arial', size: 36, bold: true, color: C.navy })],
  spacing: { before: 480, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.teal, space: 4 } },
});

const h2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text: t, font: 'Arial', size: 28, bold: true, color: C.teal })],
  spacing: { before: 320, after: 120 },
});

const h3 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text: t, font: 'Arial', size: 24, bold: true, color: C.slate })],
  spacing: { before: 200, after: 80 },
});

const p = (t, opts={}) => new Paragraph({
  children: [new TextRun({ text: t, font: 'Arial', size: 22,
    color: opts.color||C.slate, bold: opts.bold||false, italics: opts.italic||false })],
  spacing: { before: 40, after: 40 },
});

const bullet = (t, ref='bullets') => new Paragraph({
  numbering: { reference: ref, level: 0 },
  children: [new TextRun({ text: t, font: 'Arial', size: 22, color: C.slate })],
  spacing: { before: 40, after: 40 },
});

const note = (text, type='note') => {
  const cfg = {
    note:    { bg: 'EFF6FF', bc: '3B82F6', lbl: 'NOTE',    tc: '1D4ED8' },
    tip:     { bg: 'F0FDF4', bc: C.green,  lbl: 'TIP',     tc: '166534' },
    warning: { bg: 'FFFBEB', bc: C.amber,  lbl: 'WARNING', tc: '92400E' },
    danger:  { bg: 'FEF2F2', bc: C.red,    lbl: 'CAUTION', tc: '991B1B' },
  };
  const c = cfg[type]||cfg.note;
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 10, color: c.bc },
                 bottom: b(c.bc), left: { style: BorderStyle.SINGLE, size: 10, color: c.bc }, right: b(c.bc) },
      shading: { fill: c.bg, type: ShadingType.CLEAR },
      margins: { top: 90, bottom: 90, left: 160, right: 160 },
      width: { size: 9360, type: WidthType.DXA },
      children: [new Paragraph({ children: [
        new TextRun({ text: c.lbl + '  ', font: 'Arial', size: 20, bold: true, color: c.tc }),
        new TextRun({ text, font: 'Arial', size: 20, color: C.slate })],
        spacing: { before: 0, after: 0 } })],
    })]})],
  });
};

// ── Status badge cell ─────────────────────────────────────────────
const STATUS_COLORS = {
  'Pass':     { bg: 'D1FAE5', tc: '065F46' },
  'Fail':     { bg: 'FEE2E2', tc: '991B1B' },
  'Manual':   { bg: 'EFF6FF', tc: '1D4ED8' },
  'Auto':     { bg: 'F5F3FF', tc: '4C1D95' },
  'Critical': { bg: 'FFF7ED', tc: 'C2410C' },
};

const thead = (cols, widths) => new TableRow({ tableHeader: true,
  children: cols.map((col, i) => new TableCell({
    borders: ab(C.teal), shading: { fill: C.teal, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    width: { size: widths[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text: col, font: 'Arial', size: 20, bold: true, color: C.white })],
      alignment: AlignmentType.LEFT })],
  })),
});

const trow = (cells, widths, shade=false, statusCol=-1) => new TableRow({
  children: cells.map((cell, i) => {
    const s = String(cell);
    const sc = statusCol === i && STATUS_COLORS[s] ? STATUS_COLORS[s] : null;
    return new TableCell({
      borders: ab(C.border),
      shading: { fill: sc ? sc.bg : shade ? 'F8FAFC' : C.white, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 120, right: 120 },
      width: { size: widths[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
      children: [new Paragraph({ children: [new TextRun({
        text: s, font: i === 0 ? 'Courier New' : 'Arial',
        size: i === 0 ? 18 : 20, bold: sc ? true : false,
        color: sc ? sc.tc : i === 0 ? C.teal : C.slate,
      })], spacing: { before: 0, after: 0 } })],
    });
  }),
});

// ── Test case table ───────────────────────────────────────────────
// cols: [ID, Title, Precondition, Steps, Expected, SAP CLI Command, Priority]
const W = [800, 2200, 1600, 2400, 1760, 600];

const tcTable = (rows) => new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: W,
  rows: [
    thead(['TC ID', 'Test Case', 'Precondition', 'Expected Result', 'SAP CLI / API Call', 'Priority'], W),
    ...rows.map((r, i) => trow(r, W, i % 2 === 1, 5)),
  ],
});

const pLabel = (lbl) => {
  const map = { 'P0':'Critical','P1':'High','P2':'Medium','P3':'Low' };
  return map[lbl] || lbl;
};



// ═════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ═════════════════════════════════════════════════════════════════
const children = [

  // ── Cover ──
  new Paragraph({ children: [new TextRun({ text: 'VoltStartEV', font: 'Arial', size: 56, bold: true, color: C.teal })], spacing: { before: 0, after: 60 } }),
  new Paragraph({ children: [new TextRun({ text: 'End-to-End Test Cases', font: 'Arial', size: 40, bold: true, color: C.navy })], spacing: { before: 0, after: 80 } }),
  new Paragraph({ children: [
    new TextRun({ text: 'Version 1.0  ', font: 'Arial', size: 20, color: C.gray }),
    new TextRun({ text: '|  June 2026  |  ', font: 'Arial', size: 20, color: C.gray }),
    new TextRun({ text: 'Frontend · Backend · OCPP · SteVe', font: 'Arial', size: 20, color: C.gray }),
  ], spacing: { before: 0, after: 0 } }),
  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.teal, space: 8 } }, children: [], spacing: { before: 160, after: 320 } }),

  // ── Test Environment ──
  h1('Test Environment'),
  new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2400, 6960],
    rows: [
      thead(['Component', 'Details'], [2400, 6960]),
      ...([
        ['Backend URL',      'http://136.113.7.146:3000'],
        ['SteVe URL',        'http://136.113.7.146:8080'],
        ['OCPP WS',          'ws://136.113.7.146:8080/steve/websocket/CentralSystemService/{chargeBoxId}'],
        ['Test User',        'qatest001 / QATest123!  (userId=33, role=fleet_admin)'],
        ['Test User 2',      'qatest002 / QATest123!  (userId=30, role=customer)'],
        ['Test OCPP Tag',    'QATEST001 (system tag), A1A2A3A4 (external RFID)'],
        ['Simulator',        'SAP e-mobility-charging-stations-simulator (CLI)'],
        ['Test Chargers',    'CS-HPC350K-00001 (DC 350kW), CS-AC7K-00001 (AC 7kW), CS-SCHUKO3K-00001 (3-pin)'],
        ['Mobile',           'Expo Go SDK 53 · React Native · Android/iOS'],
        ['DB',               'MySQL 8.0  stevedb + voltstartev_db'],
      ]).map((r, i) => trow(r, [2400, 6960], i % 2 === 1)),
    ],
  }),
  sp(8),
  note('Run all SAP CLI commands from the simulator directory. Prefix with the station hashId obtained from: evse-cli station list', 'tip'),
  sp(8),

  // ── Priority Legend ──
  h1('Priority Legend'),
  new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 2400, 5760],
    rows: [
      thead(['Priority', 'Label', 'Description'], [1200, 2400, 5760]),
      ...([
        ['P0', 'Critical',  'App unusable if fails. Must pass before any release.'],
        ['P1', 'High',      'Core user flow broken. Must pass before QA sign-off.'],
        ['P2', 'Medium',    'Important feature degraded. Fix before release if possible.'],
        ['P3', 'Low',       'Edge case or cosmetic. Track and fix in next sprint.'],
      ]).map((r, i) => trow(r, [1200, 2400, 5760], i % 2 === 1, 1)),
    ],
  }),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 1 — AUTHENTICATION
  // ════════════════════════════════════════════════════════════════
  h1('Module 1 — Authentication'),
  h2('1.1 Registration'),
  tcTable([
    ['TC-AUTH-001', 'New user registration with valid data', 'App installed, no existing account', 'POST /api/users/register returns 201, JWT token in response, user row in users table, OCPP tag auto-generated in user_tags and stevedb.ocpp_tag', 'curl -X POST .../register -d \'{"username":"newuser","email":"new@test.com","password":"Test1234!"}\'', 'P0'],
    ['TC-AUTH-002', 'Registration with duplicate email', 'TC-AUTH-001 completed', 'API returns 409 Conflict, no duplicate user created, app shows "Email already registered" alert', 'Same request as TC-AUTH-001', 'P0'],
    ['TC-AUTH-003', 'Registration with weak password', 'None', 'API returns 400, error mentions password requirements, no user created', 'POST /register with password="abc"', 'P1'],
    ['TC-AUTH-004', 'OCPP tag auto-assigned on registration', 'TC-AUTH-001 completed', 'SELECT ocpp_tag_id FROM user_tags WHERE app_user_id=NEW_ID returns VSE-{id}-{xxxx} format, same tag exists in stevedb.ocpp_tag', 'mysql: SELECT * FROM user_tags WHERE app_user_id=NEW_ID', 'P0'],
  ]),
  sp(8),
  h2('1.2 Login'),
  tcTable([
    ['TC-AUTH-010', 'Login with valid credentials', 'User exists in DB', 'POST /api/users/login returns 200, token in response, token stored in AsyncStorage, app navigates to MapScreen', 'curl -X POST .../login -d \'{"username":"qatest001","password":"QATest123!"}\'', 'P0'],
    ['TC-AUTH-011', 'Login with wrong password', 'User exists', 'Returns 401, app shows "Invalid credentials" alert, no token stored', 'POST /login with wrong password', 'P0'],
    ['TC-AUTH-012', 'JWT token expiry handling', 'Login completed, manually expire token in DB or wait', 'API returns 401 with "expired" message, app clears AsyncStorage, navigates to Login screen', 'Modify JWT_SECRET temporarily then call any authenticated endpoint', 'P1'],
    ['TC-AUTH-013', 'Login persists across app restart', 'User logged in, close and reopen app', 'App auto-navigates to MapScreen without showing login, token loaded from AsyncStorage', 'Close Expo Go, reopen', 'P1'],
    ['TC-AUTH-014', 'Logout clears all state', 'User logged in', 'All AsyncStorage keys cleared, authStore reset, WebSocket disconnected, app shows Login screen', 'Tap Logout in Profile → confirm', 'P1'],
  ]),
  sp(8),
  h2('1.3 WebSocket Authentication'),
  tcTable([
    ['TC-WS-001', 'WebSocket connects with token in URL', 'User logged in, app open', 'Logs show "WS authenticated, userId: X", backend logs show "WebSocket connected: user_X"', 'Check mobile logs after login', 'P0'],
    ['TC-WS-002', 'WebSocket rejected with invalid token', 'Manually corrupt token', 'Backend closes WS with code 4001, mobile logs show "Authentication required", reconnect not attempted', 'Modify token in AsyncStorage to invalid value', 'P1'],
    ['TC-WS-003', 'WebSocket reconnects after network drop', 'Session active, simulate network loss', 'After network restored, WS reconnects automatically, telemetry resumes within 30s, exponential backoff visible in logs', 'Toggle airplane mode on device', 'P1'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 2 — MAP & CHARGER DISCOVERY
  // ════════════════════════════════════════════════════════════════
  h1('Module 2 — Map & Charger Discovery'),
  h2('2.1 Map Loading'),
  tcTable([
    ['TC-MAP-001', 'Map loads with all 12 chargers visible', 'Logged in, location permission granted', 'GET /api/chargers returns 12 chargers, all markers visible on map, charger count badge shows "12/12 chargers"', 'curl -H "Authorization: Bearer $TOKEN" .../api/chargers | python3 -m json.tool', 'P0'],
    ['TC-MAP-002', 'Charger config cache hit on second load', 'TC-MAP-001 run once', 'Second GET /api/chargers shows "CHARGER CONFIG CACHE HIT" and "PRICING CACHE HIT" in backend logs, response < 200ms', 'Watch backend logs during second charger load', 'P1'],
    ['TC-MAP-003', 'Map centers on user location', 'GPS enabled, location permission granted', 'Blue dot appears at device location, map animates to user position within 2 seconds of opening', 'Allow location permission, check map', 'P1'],
    ['TC-MAP-004', 'Charger marker colors by status', 'Chargers in various states', 'Available=green, Busy/Charging=amber/blue, Faulted=red, Offline=gray, Reserved=purple', 'evse-cli ocpp status-notification --connector-id 1 --status Charging {hash}', 'P1'],
    ['TC-MAP-005', 'Map markers update without 2-min poll', 'MapScreen open, charger Available', 'After evse-cli status change, map marker updates within 3 seconds via WebSocket, not waiting for poll', 'evse-cli ocpp status-notification --connector-id 1 --status Charging {hash}', 'P0'],
  ]),
  sp(8),
  h2('2.2 Charger Modal'),
  tcTable([
    ['TC-MAP-010', 'Tap marker opens charger detail modal', 'Map loaded with chargers', 'Modal opens with charger name, status, connector count, pricing, "Start Charging" button visible', 'Tap any green marker on map', 'P0'],
    ['TC-MAP-011', 'Modal shows live connector statuses', 'Charger modal open', 'Each connector shows current OCPP status (Available/Charging/etc) with correct color', 'Open modal while evse-cli sends various statuses', 'P1'],
    ['TC-MAP-012', 'Pricing estimate shows in modal', 'Charger has active pricing in charger_pricing table', 'Modal shows "₹X.XX/kWh" or tiered rates, pricing-estimate endpoint called with user vehicle data', 'Check GET .../api/chargers/{id}/pricing-estimate response', 'P1'],
    ['TC-MAP-013', 'Modal updates when connector status changes', 'Modal open, charger Available', 'When evse-cli sends Charging status, modal connector card updates from green to blue in real time without closing modal', 'evse-cli ocpp status-notification --connector-id 1 --status Charging {hash}', 'P0'],
    ['TC-MAP-014', 'Offline charger shows correct state', 'Stop the simulator for a charger', 'Charger shows gray "Offline" marker after heartbeat timeout (5 min), modal shows "Offline" status', 'evse-cli connection close {hash} then wait 5 min', 'P2'],
  ]),
  sp(8),
  h2('2.3 Filters'),
  tcTable([
    ['TC-FILTER-001', 'Filter by Available only', 'Multiple chargers with mixed statuses', 'Only Available chargers shown, count badge updates, Busy/Offline markers hidden', 'Set availability filter to "Available", verify count decreases', 'P1'],
    ['TC-FILTER-002', 'Filter by min power 50kW', 'Chargers with various max_power_w values', 'Only DC fast chargers (50kW+) shown, AC 7kW chargers hidden', 'Set minPower=50 filter', 'P1'],
    ['TC-FILTER-003', 'Filter by connector type CCS2', 'charger_capabilities table populated', 'Only chargers with CCS2 connectors shown', 'Set connectorType=CCS2 filter', 'P2'],
    ['TC-FILTER-004', 'Filter by vehicle type 2W', 'charger_capabilities with 2W entries', 'Only Bharat AC/DC and 3-Pin chargers shown', 'Set vehicleType=2W filter', 'P2'],
    ['TC-FILTER-005', 'Reset filters shows all chargers', 'Filters applied in TC-FILTER-001 through 004', 'All 12 chargers visible, count badge shows 12/12', 'Tap Reset Filters', 'P1'],
    ['TC-FILTER-006', 'Price filter uses max tier for tiered pricing', 'HPC350K has tiered pricing up to ₹28/kWh', 'Setting maxPrice=₹20 hides HPC350K (max tier ₹28), does not show it at base rate ₹16', 'Set maxPrice=20, verify HPC350K hidden', 'P1'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 3 — CHARGING SESSION
  // ════════════════════════════════════════════════════════════════
  h1('Module 3 — Charging Session'),
  h2('3.1 Session Start'),
  tcTable([
    ['TC-SESSION-001', 'Start charging via app (happy path)', 'Wallet >= ₹50, valid tag, charger Available', 'POST /api/charging/start returns 200, SteVe receives RemoteStartTransaction, evse-cli starts transaction, session row created in charging_sessions, SessionScreen appears', 'evse-cli atg start {hash} --connector-ids 1 (to auto-respond)', 'P0'],
    ['TC-SESSION-002', 'Session blocked when wallet < ₹50', 'User wallet balance = ₹30', 'POST /api/charging/start returns 402 INSUFFICIENT_BALANCE, app shows "Add Money" alert with Wallet navigation', 'Set balance: UPDATE wallets SET balance=30 WHERE user_id=33', 'P0'],
    ['TC-SESSION-003', 'Session blocked when wallet = ₹0', 'User wallet balance = ₹0', 'Same as TC-SESSION-002, balance check fires before SteVe call', 'UPDATE wallets SET balance=0 WHERE user_id=33', 'P0'],
    ['TC-SESSION-004', 'Correct idTag sent to SteVe', 'User has primary tag A1A2A3A4 set', 'Backend logs show idTag=A1A2A3A4 in StartTransaction, not QATEST001', 'Check backend logs: "Charging session start requested"', 'P0'],
    ['TC-SESSION-005', 'Session start with expired tag blocked', 'Set expiry_date in past for QATEST001 in stevedb.ocpp_tag', 'Returns 403 "Your RFID tag has expired", no SteVe call made', 'UPDATE stevedb.ocpp_tag SET expiry_date="2020-01-01" WHERE id_tag="QATEST001"', 'P1'],
    ['TC-SESSION-006', 'Pricing stored at session creation', 'Session started on CS-HPC350K-00001', 'charging_sessions.rate_per_kwh and session_fee populated at INSERT, not null', 'SELECT rate_per_kwh, session_fee FROM charging_sessions WHERE steve_transaction_pk=LATEST', 'P0'],
    ['TC-SESSION-007', 'Session fee charged only once', 'HPC350K has ₹100 session fee', 'Session fee appears in first meter value cost delta only, subsequent meter values charge only energy rate', 'Watch backend logs: "isFirstChunk" flag in cost calculation', 'P0'],
  ]),
  sp(8),
  h2('3.2 Live Telemetry'),
  tcTable([
    ['TC-TEL-001', 'Telemetry updates in SessionScreen every 5s', 'Active charging session', 'SessionScreen shows updating kWh, cost, power, current, voltage every ~5 seconds via WebSocket', 'evse-cli ocpp meter-values --connector-id 1 {hash}', 'P0'],
    ['TC-TEL-002', 'Cost calculation correct for per_kwh model', 'Session on AC7K (₹7.50/kWh)', 'After 1 kWh delivered, costSoFar = ₹7.50 ± ₹0.01', 'Send MeterValues with 1000 Wh, check telemetry:update.costSoFar', 'P0'],
    ['TC-TEL-003', 'Tiered pricing applies correct rate per interval', 'Session on HPC350K (tiered ₹16/22/28/kWh)', 'When powerW=80000W, rate=₹22/kWh. When powerW=200000W, rate=₹28/kWh', 'Send MeterValues with powerW=80000 then 200000, check backend logs "ratePerKwh"', 'P0'],
    ['TC-TEL-004', 'Cost is monotonically increasing', 'Active session, multiple meter values', 'costSoFar never decreases between consecutive telemetry updates (monotonic guard)', 'Send 10 meter values, verify each costSoFar >= previous', 'P0'],
    ['TC-TEL-005', 'Backward meter value ignored', 'Active session, meterWh > 0', 'If MeterValues reports lower Wh than previous, backend ignores it, logs "Ignoring backward meter", cost unchanged', 'Send MeterValues with meterWh lower than previous reading', 'P1'],
    ['TC-TEL-006', 'SOC shown in SessionScreen when available', 'Charger sends SoC measurand', 'Battery percentage visible in SessionScreen stats row', 'evse-cli ocpp meter-values -p \'{"meterValue":[{"sampledValue":[{"measurand":"SoC","value":"65","unit":"Percent"}]}]}\' {hash}', 'P2'],
    ['TC-TEL-007', 'Temperature shown when above threshold', 'Charger sends Temperature measurand > 45C', 'Temperature displayed in amber, > 60C shows red warning icon', 'Send MeterValues with Temperature=55, then 65', 'P2'],
    ['TC-TEL-008', 'Telemetry updates HistoryScreen live cost', 'HistoryScreen open, active session', 'Active session card shows updating cost in blue, not zero', 'Open History tab while session active', 'P0'],
    ['TC-TEL-009', 'History pull-to-refresh preserves live cost', 'TC-TEL-008 in progress', 'After pull-to-refresh, live cost does not reset to zero — shows last_cost from DB as seed', 'Pull down to refresh on HistoryScreen during active session', 'P0'],
  ]),
  sp(8),
  h2('3.3 Session Stop'),
  tcTable([
    ['TC-STOP-001', 'Stop session via app (happy path)', 'Active session', 'POST /api/charging/stop → RemoteStopTransaction sent to SteVe → StopTransaction webhook received → session.status=completed, total_cost calculated, wallet deducted', 'Tap Stop in SessionScreen', 'P0'],
    ['TC-STOP-002', 'Wallet correctly deducted on stop', 'Active session, wallet=₹200, session cost=₹45', 'wallet.balance = ₹155 after session_completed webhook processed', 'SELECT balance FROM wallets WHERE user_id=33 after stop', 'P0'],
    ['TC-STOP-003', 'Session appears in History after stop', 'TC-STOP-001 completed', 'New completed session visible in HistoryScreen with correct kWh and cost', 'Navigate to History tab after stopping', 'P0'],
    ['TC-STOP-004', 'Map marker returns to Available after stop', 'TC-STOP-001 completed', 'Map marker turns green within 5 seconds of StopTransaction — via WebSocket not poll', 'Watch map after tapping Stop', 'P0'],
    ['TC-STOP-005', 'Stop button debounce prevents double-stop', 'Active session', 'Tapping Stop multiple times quickly only fires one RemoteStopTransaction', 'Tap Stop rapidly 3 times, check backend logs for single stop request', 'P1'],
    ['TC-STOP-006', 'Session stopped via charger (local stop)', 'Active session, physically stop at charger', 'SteVe receives StopTransaction with reason=Local, webhook fires, session completes, wallet deducted correctly', 'evse-cli transaction stop --transaction-id {txId} {hash}', 'P1'],
    ['TC-STOP-007', 'Stop reason EmergencyStop handled', 'Active session', 'StopTransaction with reason=EmergencyStop creates completed session, user notified', 'evse-cli transaction stop -p \'{"reason":"EmergencyStop"}\' --transaction-id {txId} {hash}', 'P2'],
  ]),
  sp(8),
  h2('3.4 Auto-Stop'),
  tcTable([
    ['TC-AUTO-001', 'SOC auto-stop fires at target SOC', 'Session active, user target SOC=80%, real SOC < 80%', 'When SOC reaches 80%, RemoteStopTransaction sent, session_completed, push notification received', 'evse-cli ocpp meter-values -p with SoC values 70,75,79,80', 'P0'],
    ['TC-AUTO-002', 'SOC auto-stop not triggered by simulator noise', 'Random SOC values (0-100% each interval)', 'Auto-stop not triggered when SOC jump > 20% per interval (noise guard)', 'Send SoC: 20,90,10,80,5 — verify no stop fires', 'P0'],
    ['TC-AUTO-003', 'SOC auto-stop fires only once', 'Session active, target SOC=80%', 'RemoteStopTransaction sent exactly once even if multiple meter values arrive simultaneously at 80%', 'Send 3 rapid meter values all with SoC=80, check backend logs for single stop', 'P0'],
    ['TC-AUTO-004', 'Low balance auto-stop fires at ₹5 remaining', 'Wallet=₹20, session on HPC350K (₹28/kWh)', 'When costSoFar reaches ₹15 (₹5 threshold), RemoteStop sent, balance_critical WebSocket event received', 'Send MeterValues accumulating to ₹15 cost', 'P0'],
    ['TC-AUTO-005', 'Low balance auto-stop fires at exactly ₹0', 'Wallet=₹0', 'Immediate RemoteStop on first meter value, session never accumulates cost beyond balance', 'UPDATE wallets SET balance=0, then start session and send any MeterValues', 'P0'],
    ['TC-AUTO-006', 'Balance_critical alert shown in app', 'TC-AUTO-004 in progress', 'App shows "Wallet Balance Critical" alert with "Add Money" button navigating to Wallet tab', 'Watch app during low balance auto-stop', 'P1'],
    ['TC-AUTO-007', 'SOC target update mid-session works', 'Active session, target SOC=80%', 'User changes target to 95% in Profile → PUT .../target-soc → WebSocket soc_target_updated event → auto-stop now fires at 95%', 'Change target SOC in Profile while session active, send SoC=82 (should not stop), then SoC=96 (should stop)', 'P1'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 4 — WALLET
  // ════════════════════════════════════════════════════════════════
  h1('Module 4 — Wallet'),
  tcTable([
    ['TC-WALLET-001', 'Wallet balance displayed correctly', 'User has wallet with balance', 'GET /api/wallet returns correct balance, WalletScreen shows same value', 'curl -H "Authorization: Bearer $TOKEN" .../api/wallet', 'P0'],
    ['TC-WALLET-002', 'Load wallet via Razorpay WebView', 'User on WalletScreen', 'Razorpay WebView opens, payment succeeds, balance increases by loaded amount, transaction appears in wallet_transactions', 'Tap "Add Money" → complete test payment with Razorpay test card', 'P0'],
    ['TC-WALLET-003', 'Razorpay webhook signature verified', 'Valid Razorpay webhook with correct secret', 'Wallet credited only when HMAC-SHA256 signature matches, rejected otherwise', 'Send webhook with correct then incorrect signature', 'P0'],
    ['TC-WALLET-004', 'Transaction history shows all types', 'Multiple transactions: credit, debit, refund', 'WalletScreen shows all transactions with correct icons, amounts, dates', 'Complete a session (debit) and load money (credit)', 'P1'],
    ['TC-WALLET-005', 'Live balance shown during session', 'Active charging session, SessionScreen open', 'Wallet balance card shows decreasing balance in real time (refreshed every 30s)', 'Watch SessionScreen during active session', 'P1'],
    ['TC-WALLET-006', 'Duplicate payment order rejected', 'Valid payment order created', 'Replaying the same Razorpay webhook does not credit wallet twice (idempotency check)', 'POST same webhook payload twice', 'P0'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 5 — RFID & TAGS
  // ════════════════════════════════════════════════════════════════
  h1('Module 5 — RFID & Charging Tags'),
  tcTable([
    ['TC-RFID-001', 'Add external RFID card', 'User on Profile → RFID section', 'POST /api/users/me/rfid with valid UID → row in user_tags, row in stevedb.ocpp_tag, row in stevedb.user_ocpp_tag, card visible in RFID section', 'Tap + Add RFID → enter A1A2A3A4', 'P0'],
    ['TC-RFID-002', 'Set external RFID as primary', 'TC-RFID-001 completed', 'PUT .../primary → is_primary=1 for new card, is_primary=0 for previous primary, OCPP tag in Profile header updates', 'Tap Set Primary on new card', 'P0'],
    ['TC-RFID-003', 'Remove RFID card (soft delete)', 'External RFID card exists, is not primary', 'DELETE .../rfid/{id} → user_tags.is_active=0, system tag promoted to primary, SteVe tag blocked (expiry=2000-01-01)', 'Tap Remove on RFID card', 'P0'],
    ['TC-RFID-004', 'Cannot remove system tag', 'System tag (App Tag) visible in RFID section', 'Tap Remove on system tag → alert "Cannot Remove — used as fallback", no API call made', 'Attempt to remove App Tag', 'P1'],
    ['TC-RFID-005', 'RFID tag validates for session start', 'External RFID A1A2A3A4 is primary', 'POST /api/charging/start with idTag=A1A2A3A4 → validateTagForUser checks user_tags table → passes', 'Start session with external RFID as primary tag', 'P0'],
    ['TC-RFID-006', 'Profile OCPP tag updates after set primary', 'Two cards: QATEST001 and A1A2A3A4', 'After setting A1A2A3A4 as primary, Profile Account Info row shows A1A2A3A4 not QATEST001', 'Change primary, check Profile header', 'P1'],
    ['TC-RFID-007', 'Steve sync queue retries on failure', 'SteVe down, RFID deleted', 'Deletion queued in steve_sync_queue table, processed when SteVe comes back, exponential backoff in logs', 'Stop SteVe, delete RFID, restart SteVe, verify sync completes', 'P1'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 6 — RESERVATIONS
  // ════════════════════════════════════════════════════════════════
  h1('Module 6 — Reservations'),
  tcTable([
    ['TC-RES-001', 'Reserve a connector', 'Charger Available, no active reservation', 'POST /api/reservations → SteVe ReserveNow OCPP sent, connector shows Reserved in modal, app_reservations row created', 'Tap Reserve in charger modal', 'P1'],
    ['TC-RES-002', 'Cannot start session on connector reserved by another user', 'Connector reserved by qatest002', 'Attempt to start session as qatest001 → blocked, connector shows Reserved badge', 'Reserve as user2, attempt start as user1', 'P1'],
    ['TC-RES-003', 'Cancel reservation', 'Active reservation exists', 'DELETE /api/reservations/{id} → CancelReservation sent to SteVe → connector returns to Available', 'Tap Cancel Reservation in modal', 'P1'],
    ['TC-RES-004', 'Active reservation shown in modal on open', 'User has active reservation', 'Opening charger modal shows "Your Reservation" badge on reserved connector', 'Have reservation, open modal', 'P2'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 7 — VEHICLES & PROFILE
  // ════════════════════════════════════════════════════════════════
  h1('Module 7 — Vehicles & Profile'),
  tcTable([
    ['TC-VEH-001', 'Add vehicle', 'User on Profile', 'POST /api/users/me/vehicles → vehicle row in user_vehicles, appears in Profile vehicles section', 'Add vehicle: MG ZS EV, 50.3 kWh', 'P1'],
    ['TC-VEH-002', 'Set primary vehicle', 'Two vehicles exist', 'PUT .../primary → is_primary=1 for selected, 0 for others', 'Tap Set Primary on second vehicle', 'P1'],
    ['TC-VEH-003', 'Target SOC control always visible', 'Vehicle exists', 'TargetSocControl renders below primary vehicle card, shows current SOC target, presets visible', 'Check Profile screen', 'P1'],
    ['TC-VEH-004', 'Target SOC update persists', 'Vehicle with target_soc=80', 'Change to 90% via TargetSocControl → PUT .../target-soc → DB updated → pricing estimate uses new SOC', 'Change target SOC, verify DB: SELECT target_soc FROM user_vehicles WHERE is_primary=1', 'P1'],
    ['TC-VEH-005', 'Target SOC update during active session', 'Active charging session, target=80%', 'Change target to 95% → PUT .../target-soc → soc_target_updated WebSocket event received → SessionScreen target updates', 'Change target during active session, watch SessionScreen', 'P1'],
    ['TC-VEH-006', 'Delete vehicle (not primary)', 'Two vehicles, primary and secondary', 'DELETE .../vehicles/{id} → vehicle removed from DB, not visible in Profile', 'Tap Remove on non-primary vehicle', 'P1'],
    ['TC-VEH-007', 'Cannot delete primary vehicle when it is the only vehicle', 'Only one vehicle, is_primary=1', 'Deletion blocked with appropriate message', 'Attempt to remove only vehicle', 'P2'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 8 — FLEET MANAGEMENT
  // ════════════════════════════════════════════════════════════════
  h1('Module 8 — Fleet Management'),
  p('Fleet features are gated behind FLEET_ENABLED=true env flag and role=fleet_admin JWT claim.'),
  sp(4),
  tcTable([
    ['TC-FLEET-001', 'Fleet section visible for fleet_admin', 'qatest001 has role=fleet_admin', 'FleetSection renders in Profile, "Create Fleet" or fleet details visible', 'Login as qatest001, check Profile', 'P1'],
    ['TC-FLEET-002', 'Fleet section hidden for regular user', 'qatest002 has role=customer', 'FleetSection not rendered in Profile for customer role', 'Login as qatest002, check Profile', 'P1'],
    ['TC-FLEET-003', 'Create fleet', 'fleet_admin user, no existing fleet', 'POST /api/fleet → fleet row in fleets table, user assigned as admin in fleet_members', 'Create fleet with name "Test Fleet"', 'P2'],
    ['TC-FLEET-004', 'Fleet billing mode: fleet pays', 'Active fleet with billing_mode=fleet_pays', 'Session started by fleet member billed to fleet wallet, not personal wallet', 'Start session as fleet member', 'P2'],
    ['TC-FLEET-005', 'Fleet monthly limit enforcement', 'Fleet has monthly_limit=₹1000, spent ₹999', 'Next session start blocked when monthly limit reached', 'Set monthly_limit low and attempt session start', 'P2'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 9 — PUSH NOTIFICATIONS
  // ════════════════════════════════════════════════════════════════
  h1('Module 9 — Push Notifications'),
  p('Note: Push notifications require a real dev build (not Expo Go). Test with dev build or verify via backend logs.'),
  sp(4),
  tcTable([
    ['TC-PUSH-001', 'Session started push notification', 'Push token registered, session starts', 'Push notification received: "⚡ Charging Started! {chargeBoxId} · Connector {connectorId}"', 'Start session, check device notifications', 'P1'],
    ['TC-PUSH-002', 'Session completed push notification', 'Active session completes', 'Push notification received with final cost and energy', 'Stop session, check device notifications', 'P1'],
    ['TC-PUSH-003', 'SOC target reached push notification', 'TC-AUTO-001 completed', 'Push notification: "Target Charge Reached — vehicle at X%"', 'Check device after SOC auto-stop', 'P1'],
    ['TC-PUSH-004', 'Low balance push notification', 'TC-AUTO-004 completed', 'Push notification: "Low Balance — Stopping Charge ₹X.XX remaining"', 'Check device after balance auto-stop', 'P1'],
    ['TC-PUSH-005', 'Push token persists across login', 'User logs out and back in', 'push_token column in users table updated on re-login, notifications still received', 'Logout, login, start session', 'P2'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 10 — OCPP 1.6J COMPLIANCE
  // ════════════════════════════════════════════════════════════════
  h1('Module 10 — OCPP 1.6J Compliance'),
  tcTable([
    ['TC-OCPP-001', 'BootNotification accepted', 'Charger reconnects', 'SteVe responds Accepted, registration_status=Accepted in stevedb, charger shows online in app within 2 min', 'evse-cli ocpp boot-notification {hash}', 'P0'],
    ['TC-OCPP-002', 'Heartbeat keeps charger online', 'Charger connected', 'last_heartbeat_timestamp updates every 60s in stevedb.charge_box', 'evse-cli ocpp heartbeat {hash} (repeat)', 'P1'],
    ['TC-OCPP-003', 'StatusNotification updates DB immediately', 'Charger changes state', 'connector_status row inserted in stevedb, WebSocket broadcast fired, app map updates within 3s', 'evse-cli ocpp status-notification --connector-id 1 --status Charging --error-code NoError {hash}', 'P0'],
    ['TC-OCPP-004', 'RemoteStartTransaction response', 'Charger connected, Available', 'SteVe receives RemoteStart, returns Accepted, simulator starts transaction', 'POST /api/charging/start, watch SteVe logs', 'P0'],
    ['TC-OCPP-005', 'RemoteStopTransaction response', 'Active transaction', 'SteVe receives RemoteStop, returns Accepted, simulator stops transaction', 'POST /api/charging/stop, watch SteVe logs', 'P0'],
    ['TC-OCPP-006', 'SetChargingProfile accepted', 'Charger connected, operator role', 'POST /api/chargers/{id}/connectors/1/charging-profile → SteVe SetChargingProfile sent, charger limits power', 'curl -X POST .../charging-profile -d \'{"maxPowerWatts":50000}\'', 'P1'],
    ['TC-OCPP-007', 'ClearChargingProfile restores default', 'Charging profile applied', 'DELETE /api/chargers/{id}/connectors/1/charging-profile → ClearChargingProfile sent, charger returns to full power', 'curl -X DELETE .../charging-profile', 'P1'],
    ['TC-OCPP-008', 'ChangeAvailability Inoperative', 'Charger connected, operator role', 'PUT /api/chargers/{id}/availability with Inoperative → OCPP ChangeAvailability sent, connector shows Unavailable', 'curl -X PUT .../availability -d \'{"type":"Inoperative","connectorId":1}\'', 'P1'],
    ['TC-OCPP-009', 'MeterValues with all measurands extracted', 'Active session, rich meter values', 'telemetry_extractor correctly parses Energy.Active.Import.Register, Power.Active.Import, Current.Import, Voltage, SoC, Temperature', 'Send full MeterValues payload with all measurands', 'P0'],
    ['TC-OCPP-010', 'Authorization tag expiry checked', 'OCPP tag has past expiry_date', 'Session start returns 403 "tag expired", no RemoteStartTransaction sent', 'Set expiry in past for tag, attempt start', 'P1'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 11 — PERFORMANCE & RELIABILITY
  // ════════════════════════════════════════════════════════════════
  h1('Module 11 — Performance & Reliability'),
  tcTable([
    ['TC-PERF-001', 'Two users charging simultaneously', 'Two sessions active (qatest001 + qatest002)', 'Telemetry isolated per user, no cost cross-contamination, both sessions update independently', 'Start sessions as both users on different chargers simultaneously', 'P0'],
    ['TC-PERF-002', 'Charger list loads in < 500ms (cache hit)', 'Cache warmed from previous request', 'GET /api/chargers with cache hit responds in < 500ms', 'Second request after cache warms, check response time', 'P1'],
    ['TC-PERF-003', 'No double-fetch on app load', 'App cold start', 'Only one GET /api/chargers and one GET /api/charging/session/active in logs within first 3 seconds', 'Check backend logs on app launch', 'P1'],
    ['TC-PERF-004', 'Reconciliation fixes orphaned sessions', 'Active session, backend restarted mid-session', 'Reconciliation job (runs every 10 min) marks stale sessions as completed, final cost calculated from last_cost', 'pm2 restart voltstartev mid-session, wait 10 min', 'P1'],
    ['TC-PERF-005', 'Pool exhaustion handled gracefully', 'Multiple concurrent users (5+)', 'No ETIMEDOUT errors, requests queued up to queueLimit, slow queries logged as warnings', 'Start 5 simultaneous API requests', 'P1'],
    ['TC-PERF-006', 'DB atomic OTP verification (no race condition)', 'OTP verification endpoint', 'Simultaneous duplicate OTP submissions: only first succeeds, second returns "OTP already used"', 'Send same OTP twice within 100ms', 'P0'],
    ['TC-PERF-007', 'Auto-stop fires once even with concurrent meter values', 'Active session approaching SOC target', 'Multiple meter values arriving simultaneously only trigger one RemoteStop', 'Send 3 rapid meter values all at target SOC', 'P0'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // MODULE 12 — UI / UX
  // ════════════════════════════════════════════════════════════════
  h1('Module 12 — UI / UX'),
  tcTable([
    ['TC-UI-001', 'Icons render correctly on Android (MIUI)', 'Xiaomi/Redmi device', 'No emoji boxes, all BoltIcon/AppIcon components render as expected, no "Text strings must be in Text" error', 'Run on Xiaomi device, check Profile and Map screens', 'P1'],
    ['TC-UI-002', 'Map renders with tiles (not white)', 'Expo Go SDK 53', 'Map tiles visible, markers positioned correctly, user location dot shown', 'Open app on device with GPS', 'P0'],
    ['TC-UI-003', 'History screen shows zero after pull-to-refresh', 'Active session in progress', 'Pull-to-refresh on HistoryScreen preserves liveCost (does not reset to zero)', 'Pull refresh during active session', 'P0'],
    ['TC-UI-004', 'New session appears in History immediately', 'HistoryScreen open when session starts', 'session_started WebSocket event triggers fetchHistory, new active session card appears without manual refresh', 'Open History, start session on another tab', 'P0'],
    ['TC-UI-005', 'Session screen navigated automatically on start', 'Session starts via app', 'App navigates to Session tab automatically when session_started WebSocket event received', 'Start session while on Map tab', 'P1'],
    ['TC-UI-006', 'Require cycles warning suppressed', 'App running', 'Circular dependency between api.ts and authStore.ts resolved via lazy require, WARN disappears from logs', 'Check Expo logs for "Require cycle" warning', 'P2'],
    ['TC-UI-007', 'expo-notifications deprecation warning', 'App running', 'shouldShowAlert replaced with shouldShowBanner + shouldShowList in notification handler', 'Check Expo logs for deprecation warning', 'P3'],
    ['TC-UI-008', 'Connector type badge visible in modal', 'charger_capabilities table populated', 'Connector cards in modal show "CCS2", "Type2", "BharatAC" badges', 'Open charger modal, check connector cards', 'P2'],
    ['TC-UI-009', 'SOC auto-stop alert shown on correct screen', 'Session auto-stops due to SOC', 'Alert appears regardless of which tab user is on when soc_target_reached event fires', 'Be on History tab when SOC auto-stop fires', 'P1'],
    ['TC-UI-010', 'Wallet low balance warning during session', 'Balance < ₹50 during active session', 'SessionScreen balance card shown in amber with warning indicator', 'Set balance to ₹30 then start session', 'P1'],
  ]),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // SAP CLI TEST SCRIPTS
  // ════════════════════════════════════════════════════════════════
  h1('Appendix A — SAP CLI Quick Reference'),
  p('Run all commands from the SAP simulator directory. Get {hash} from: evse-cli station list'),
  sp(8),
  new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2400, 6960],
    rows: [
      thead(['Action', 'SAP CLI Command'], [2400, 6960]),
      ...([
        ['List all stations',             'evse-cli station list'],
        ['Check simulator state',         'evse-cli simulator state'],
        ['Set connector Available',       'evse-cli ocpp status-notification --connector-id 1 --error-code NoError --status Available {hash}'],
        ['Set connector Charging',        'evse-cli ocpp status-notification --connector-id 1 --error-code NoError --status Charging {hash}'],
        ['Set connector Faulted',         'evse-cli ocpp status-notification --connector-id 1 --error-code GroundFailure --status Faulted {hash}'],
        ['Send Heartbeat',                'evse-cli ocpp heartbeat {hash}'],
        ['Send BootNotification',         'evse-cli ocpp boot-notification {hash}'],
        ['Start ATG (auto transactions)', 'evse-cli atg start {hash} --connector-ids 1'],
        ['Stop ATG',                      'evse-cli atg stop {hash}'],
        ['Manual transaction start',      'evse-cli transaction start --connector-id 1 --id-tag QATEST001 {hash}'],
        ['Manual transaction stop',       'evse-cli transaction stop --transaction-id {txId} {hash}'],
        ['Send MeterValues (custom)',     'evse-cli ocpp meter-values --connector-id 1 -p \'{"meterValue":[{"sampledValue":[{"measurand":"Energy.Active.Import.Register","value":"1000","unit":"Wh"},{"measurand":"SoC","value":"75","unit":"Percent"}]}]}\' {hash}'],
        ['Send specific SOC value',       'evse-cli ocpp meter-values --connector-id 1 -p \'{"meterValue":[{"sampledValue":[{"measurand":"SoC","value":"80","unit":"Percent"}]}]}\' {hash}'],
        ['Open WebSocket connection',     'evse-cli connection open {hash}'],
        ['Close WebSocket connection',    'evse-cli connection close {hash}'],
      ]).map((r, i) => trow(r, [2400, 6960], i % 2 === 1)),
    ],
  }),
  sp(16),

  // ════════════════════════════════════════════════════════════════
  // DB VERIFICATION QUERIES
  // ════════════════════════════════════════════════════════════════
  h1('Appendix B — DB Verification Queries'),
  new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2800, 6560],
    rows: [
      thead(['What to Check', 'Query'], [2800, 6560]),
      ...([
        ['Active sessions',               'SELECT session_id, charge_box_id, status, last_cost, app_user_id FROM charging_sessions WHERE end_time IS NULL'],
        ['Latest session cost',           'SELECT session_id, total_cost, last_cost, energy_kwh, stop_reason FROM charging_sessions ORDER BY start_time DESC LIMIT 5'],
        ['Wallet balance after session',  'SELECT w.balance, w.user_id FROM wallets w WHERE w.user_id = 33'],
        ['User RFID tags',                'SELECT id, ocpp_tag_id, tag_type, is_primary, is_active FROM user_tags WHERE app_user_id = 33'],
        ['SteVe tag activity',            'SELECT id_tag, active_transaction_count, in_transaction FROM stevedb.ocpp_tag_activity WHERE id_tag IN (SELECT id_tag FROM user_tags WHERE app_user_id=33)'],
        ['Charger connector status',      'SELECT c.charge_box_id, c.connector_id, cs.status, cs.status_timestamp FROM stevedb.connector c JOIN stevedb.connector_status cs ON cs.connector_pk = c.connector_pk WHERE c.charge_box_id = "CS-HPC350K-00001" AND c.connector_id > 0 ORDER BY cs.status_timestamp DESC LIMIT 6'],
        ['Active pricing for all chargers','SELECT charge_box_id, pricing_model, rate_per_kwh, session_fee, display_name FROM charger_pricing WHERE is_active=1 ORDER BY charge_box_id'],
        ['Steve sync queue status',       'SELECT action, status, attempts, last_error FROM steve_sync_queue ORDER BY created_at DESC LIMIT 10'],
        ['Webhook events (last 10)',      'SELECT event_type, charge_box_id, processed_at FROM webhook_events ORDER BY processed_at DESC LIMIT 10'],
        ['Charger online status',         'SELECT charge_box_id, last_heartbeat_timestamp, registration_status FROM stevedb.charge_box ORDER BY charge_box_id'],
      ]).map((r, i) => trow(r, [2800, 6560], i % 2 === 1)),
    ],
  }),
  sp(16),

  // ── Summary stats ──
  h1('Test Coverage Summary'),
  new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [3360, 1800, 1800, 1200, 1200],
    rows: [
      thead(['Module', 'Total TCs', 'P0/P1', 'P2/P3', 'Type'], [3360, 1800, 1800, 1200, 1200]),
      ...([
        ['1 — Authentication',        '8',  '8',  '0',  'Auto + Manual'],
        ['2 — Map & Discovery',       '14', '10', '4',  'Auto + Manual'],
        ['3 — Charging Session',      '21', '18', '3',  'Auto + SAP CLI'],
        ['4 — Wallet',                '6',  '5',  '1',  'Auto + Manual'],
        ['5 — RFID & Tags',           '7',  '6',  '1',  'Auto + Manual'],
        ['6 — Reservations',          '4',  '3',  '1',  'Manual + CLI'],
        ['7 — Vehicles & Profile',    '7',  '5',  '2',  'Manual'],
        ['8 — Fleet Management',      '5',  '3',  '2',  'Manual'],
        ['9 — Push Notifications',    '5',  '4',  '1',  'Dev Build'],
        ['10 — OCPP 1.6J Compliance', '10', '8',  '2',  'Auto + CLI'],
        ['11 — Performance',          '7',  '6',  '1',  'Auto + Load'],
        ['12 — UI / UX',              '10', '7',  '3',  'Manual'],
        ['TOTAL',                     '104','83', '21', ''],
      ]).map((r, i) => trow(r, [3360, 1800, 1800, 1200, 1200], i % 2 === 1)),
    ],
  }),
  sp(8),
  note('P0/P1 tests must all pass before any production release. P2/P3 are tracked but non-blocking.', 'warning'),
];

// ── Build document ────────────────────────────────────────────────
const doc = new Document({
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]},
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 480, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial' },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ],
  },
  sections: [{
    properties: {
      page: { size: { width: 15840, height: 12240 },
              margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
              orientation: 'landscape' }
    },
    headers: { default: makeHeader() },
    footers: { default: makeFooter() },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('C:/voltstartEV/VoltStartEV_E2E_TestCases.docx', buf);
  console.log('Done — VoltStartEV_E2E_TestCases.docx written');
}).catch(err => { console.error(err); process.exit(1); });
