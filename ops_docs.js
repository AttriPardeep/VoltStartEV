const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, PageBreak,
  TabStopType, TabStopPosition, ExternalHyperlink,
} = require('docx');
const fs = require('fs');

// ── Brand colours ─────────────────────────────────────
const TEAL    = '0D9488';
const NAVY    = '0A1628';
const SLATE   = '1E293B';
const AMBER   = 'D97706';
const GREEN   = '059669';
const RED     = 'DC2626';
const LGRAY   = 'F1F5F9';
const MGRAY   = 'CBD5E1';
const DGRAY   = '475569';
const WHITE   = 'FFFFFF';
const BLACK   = '0F172A';

// ── DXA helpers ───────────────────────────────────────
const IN = n => Math.round(n * 1440);   // inches to DXA
const PW = 9360;                         // content width (US Letter, 1" margins)

// ── Common border ─────────────────────────────────────
const b = (color = MGRAY, size = 4) => ({
  style: BorderStyle.SINGLE, size, color,
});
const cellBorder = (color = MGRAY) => ({
  top: b(color), bottom: b(color), left: b(color), right: b(color),
});
const noBorder = {
  top:    { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left:   { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right:  { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};

// ── Text helpers ──────────────────────────────────────
const run = (text, opts = {}) => new TextRun({ text, font: 'Arial', ...opts });
const bold = (text, opts = {}) => run(text, { bold: true, ...opts });
const mono = (text, opts = {}) => new TextRun({ text, font: 'Courier New', size: 18, ...opts });

const para = (children, opts = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  spacing: { after: 120 },
  ...opts,
});
const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [bold(text, { size: 28, color: NAVY })],
  spacing: { before: 360, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 4 } },
});
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [bold(text, { size: 24, color: SLATE })],
  spacing: { before: 280, after: 140 },
});
const h3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [bold(text, { size: 20, color: DGRAY })],
  spacing: { before: 200, after: 100 },
});

const bullet = (text, ref = 'bul') => new Paragraph({
  numbering: { reference: ref, level: 0 },
  children: typeof text === 'string' ? [run(text, { size: 20 })] : text,
  spacing: { after: 80 },
});
const numbered = (text) => new Paragraph({
  numbering: { reference: 'num', level: 0 },
  children: typeof text === 'string' ? [run(text, { size: 20 })] : text,
  spacing: { after: 100 },
});

const spacer = (n = 160) => new Paragraph({ children: [run('')], spacing: { after: n } });

// ── Coloured label strip ──────────────────────────────
const labelStrip = (text, bg = TEAL, fg = WHITE) => new Paragraph({
  children: [bold(text, { color: fg, size: 20, allCaps: true })],
  shading: { fill: bg, type: ShadingType.CLEAR },
  spacing: { before: 200, after: 120 },
  indent: { left: IN(0.1), right: IN(0.1) },
});

// ── Info box ──────────────────────────────────────────
const infoBox = (lines, bg = 'EFF6FF', borderColor = '3B82F6') =>
  new Table({
    width: { size: PW, type: WidthType.DXA },
    columnWidths: [PW],
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top:    b(borderColor, 8),
          bottom: b(borderColor, 4),
          left:   b(borderColor, 16),
          right:  b(borderColor, 4),
        },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 160, right: 160 },
        width: { size: PW, type: WidthType.DXA },
        children: lines.map(l => para(
          typeof l === 'string' ? [run(l, { size: 20 })] : l,
          { spacing: { after: 60 } }
        )),
      })],
    })],
  });

const warningBox = (lines) => infoBox(lines, 'FFFBEB', AMBER);
const dangerBox  = (lines) => infoBox(lines, 'FEF2F2', RED);
const successBox = (lines) => infoBox(lines, 'F0FDF4', GREEN);

// ── Code block ───────────────────────────────────────
const codeBlock = (lines) => new Table({
  width: { size: PW, type: WidthType.DXA },
  columnWidths: [PW],
  rows: [new TableRow({
    children: [new TableCell({
      borders: cellBorder('1E293B'),
      shading: { fill: '0F172A', type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 180, right: 180 },
      width: { size: PW, type: WidthType.DXA },
      children: lines.map(l => new Paragraph({
        children: [new TextRun({ text: l, font: 'Courier New', size: 18, color: '22D3EE' })],
        spacing: { after: 40 },
      })),
    })],
  })],
});

// ── Two-column table row helper ───────────────────────
const W1 = Math.round(PW * 0.35);
const W2 = PW - W1;

const twoCol = (label, value, labelBg = LGRAY, highlight = false) =>
  new TableRow({
    children: [
      new TableCell({
        borders: cellBorder(),
        shading: { fill: labelBg, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: W1, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [para([bold(label, { size: 19, color: SLATE })], { spacing: { after: 0 } })],
      }),
      new TableCell({
        borders: cellBorder(),
        shading: { fill: highlight ? 'ECFDF5' : WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: W2, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [para(
          typeof value === 'string' ? [run(value, { size: 19 })] : value,
          { spacing: { after: 0 } }
        )],
      }),
    ],
  });

const tableHeader = (col1, col2) =>
  new TableRow({
    tableHeader: true,
    children: [
      new TableCell({
        borders: cellBorder(TEAL),
        shading: { fill: TEAL, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        width: { size: W1, type: WidthType.DXA },
        children: [para([bold(col1, { size: 20, color: WHITE })], { spacing: { after: 0 } })],
      }),
      new TableCell({
        borders: cellBorder(TEAL),
        shading: { fill: TEAL, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        width: { size: W2, type: WidthType.DXA },
        children: [para([bold(col2, { size: 20, color: WHITE })], { spacing: { after: 0 } })],
      }),
    ],
  });

const twoColTable = (rows) => new Table({
  width: { size: PW, type: WidthType.DXA },
  columnWidths: [W1, W2],
  rows,
});

// ── Wide 3-col table ──────────────────────────────────
const w3 = [Math.round(PW*0.28), Math.round(PW*0.22), PW - Math.round(PW*0.28) - Math.round(PW*0.22)];
const row3 = (a, b_, c, shade = WHITE) => new TableRow({
  children: [a, b_, c].map((txt, i) => new TableCell({
    borders: cellBorder(),
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    width: { size: w3[i], type: WidthType.DXA },
    children: [para(
      typeof txt === 'string' ? [run(txt, { size: 18 })] : txt,
      { spacing: { after: 0 } }
    )],
  })),
});
const head3 = (a, b_, c) => new TableRow({
  tableHeader: true,
  children: [a, b_, c].map((txt, i) => new TableCell({
    borders: cellBorder(TEAL),
    shading: { fill: TEAL, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 100, right: 100 },
    width: { size: w3[i], type: WidthType.DXA },
    children: [para([bold(txt, { size: 19, color: WHITE })], { spacing: { after: 0 } })],
  })),
});
const table3 = (rows) => new Table({
  width: { size: PW, type: WidthType.DXA },
  columnWidths: w3,
  rows,
});

// ── Step box ──────────────────────────────────────────
const stepW = [IN(0.5), PW - IN(0.5)];
const stepBox = (n, title, children) => new Table({
  width: { size: PW, type: WidthType.DXA },
  columnWidths: stepW,
  rows: [new TableRow({
    children: [
      new TableCell({
        borders: noBorder,
        shading: { fill: TEAL, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 60, right: 60 },
        width: { size: stepW[0], type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        children: [para([bold(String(n), { size: 24, color: WHITE })],
          { alignment: AlignmentType.CENTER, spacing: { after: 0 } })],
      }),
      new TableCell({
        borders: {
          top: b(MGRAY), bottom: b(MGRAY), left: b(TEAL, 12), right: b(MGRAY),
        },
        shading: { fill: LGRAY, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 120 },
        width: { size: stepW[1], type: WidthType.DXA },
        children: [
          para([bold(title, { size: 21, color: NAVY })], { spacing: { after: 80 } }),
          ...children,
        ],
      }),
    ],
  })],
});

// ── Page header / footer ──────────────────────────────
const makeHeader = (title) => new Header({
  children: [new Table({
    width: { size: PW, type: WidthType.DXA },
    columnWidths: [Math.round(PW*0.55), Math.round(PW*0.45)],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: { ...noBorder, bottom: b(TEAL, 8) },
          margins: { bottom: 80 },
          width: { size: Math.round(PW*0.55), type: WidthType.DXA },
          children: [para([
            bold('VoltStartEV', { size: 18, color: TEAL }),
            run('  |  ' + title, { size: 18, color: DGRAY }),
          ], { spacing: { after: 0 } })],
        }),
        new TableCell({
          borders: { ...noBorder, bottom: b(TEAL, 8) },
          margins: { bottom: 80 },
          width: { size: Math.round(PW*0.45), type: WidthType.DXA },
          children: [para([
            run('Operations Manual  v1.0', { size: 16, color: MGRAY }),
          ], { alignment: AlignmentType.RIGHT, spacing: { after: 0 } })],
        }),
      ],
    })],
  })],
});

const makeFooter = () => new Footer({
  children: [new Paragraph({
    children: [
      run('VoltStartEV Confidential  |  Page ', { size: 16, color: DGRAY }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, color: DGRAY }),
      run(' of ', { size: 16, color: DGRAY }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: DGRAY }),
      run('  |  Do not distribute without authorisation', { size: 16, color: MGRAY }),
    ],
    alignment: AlignmentType.CENTER,
    border: { top: b(MGRAY, 4) },
  })],
});

// ════════════════════════════════════════════════════════════════
// DOCUMENT 1 — Add a New Charger
// ════════════════════════════════════════════════════════════════
function buildAddChargerDoc() {
  const children = [

    // ── Cover ─────────────────────────────────────────
    new Paragraph({
      children: [bold('VoltStartEV', { size: 72, color: TEAL })],
      alignment: AlignmentType.CENTER,
      spacing: { before: IN(1), after: 80 },
    }),
    new Paragraph({
      children: [bold('Operations Manual', { size: 36, color: NAVY })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [bold('Adding a New Charger', { size: 48, color: SLATE })],
      alignment: AlignmentType.CENTER,
      border: {
        top:    { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 8 },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 8 },
      },
      spacing: { before: 200, after: 200 },
    }),
    new Paragraph({
      children: [
        run('Version 1.0   |   ', { size: 20, color: DGRAY }),
        run('May 2026', { size: 20, color: DGRAY }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { after: IN(1) },
    }),
    new Paragraph({ children: [new PageBreak()] }),

    // ── Overview ──────────────────────────────────────
    h1('1.  Overview'),
    para([run(
      'This document describes the complete end-to-end procedure for provisioning a new EV charger ' +
      'into the VoltStartEV network. Every step must be completed in sequence. Skipping any step ' +
      'will result in the charger appearing on the map but being unable to start sessions.',
      { size: 20 }
    )]),
    spacer(),

    infoBox([
      [bold('Estimated time: ', { size: 20 }), run('30–45 minutes per charger', { size: 20 })],
      [bold('Who can do this: ', { size: 20 }), run('VoltStartEV backend administrator only', { size: 20 })],
      [bold('Requires: ', { size: 20 }), run('SSH access to GCP server  |  MySQL access  |  SteVe admin panel', { size: 20 })],
    ]),
    spacer(),

    h2('1.1  What gets configured'),
    table3([
      head3('System', 'What is configured', 'Why'),
      row3('SteVe OCPP', 'charge_box record + connectors', 'Charger can connect via OCPP 1.6'),
      row3('App DB', 'charger_capabilities', 'Power, connector type, vehicle category'),
      row3('App DB', 'charger_pricing', 'Rate per kWh, session fee, pricing model'),
      row3('App DB', 'charger_config (legacy)', 'Max power fallback for older API paths'),
      row3('Mobile', 'Auto-reflects via GET /api/chargers', 'No app update needed'),
    ]),
    spacer(200),

    // ── Prerequisites ─────────────────────────────────
    h1('2.  Before You Start'),
    h2('2.1  Information to collect'),
    para([run(
      'Gather the following from the charger operator / site before touching any system:',
      { size: 20 }
    )]),
    spacer(80),

    twoColTable([
      tableHeader('Field', 'Example / Notes'),
      twoCol('Charge Box ID',       'CS-AC22K-PUNE-001  (see naming convention §2.2)'),
      twoCol('Physical location',   '15, Industrial Estate Rd, Hadapsar, Pune 411028'),
      twoCol('GPS coordinates',     '18.5088, 73.9259  (use Google Maps pin)'),
      twoCol('Number of connectors','2  (each connector = separate OCPP connector_id)'),
      twoCol('Connector types',     'CCS2, Type2, CHAdeMO, GB/T, BharatAC, BharatDC, ThreePin'),
      twoCol('Max power per conn.', '50 kW  (from charger spec sheet)'),
      twoCol('AC or DC',            'DC  (AC = Type2 / BharatAC typically)'),
      twoCol('Vehicle categories',  '4W  /  2W,3W  /  4W,Bus'),
      twoCol('Pricing model',       'per_kwh / tiered_power / time_of_use / per_minute'),
      twoCol('Rate',                '₹14.00 / kWh   +  ₹25 session fee'),
      twoCol('Operator contact',    'Name + phone for fault escalation'),
    ]),
    spacer(),

    h2('2.2  Charge Box ID naming convention'),
    para([run(
      'Follow this format strictly — IDs must be unique across the entire network:',
      { size: 20 }
    )]),
    spacer(80),

    codeBlock([
      'Format:  CS-{TYPE}{POWER}-{CITY}-{SEQ}',
      '',
      'Examples:',
      '  CS-AC7K-MUM-001    AC 7 kW   Mumbai charger #1',
      '  CS-AC22K-PUN-001   AC 22 kW  Pune charger #1',
      '  CS-DC50K-BLR-001   DC 50 kW  Bangalore charger #1',
      '  CS-HPC350K-DEL-001 HPC 350kW Delhi charger #1',
      '',
      'Type codes:  AC7K  AC11K  AC22K  DC50K  DC150K  HPC350K  SCHUKO',
      'City codes:  MUM  PUN  BLR  DEL  HYD  CHE  KOL  AHM',
    ]),
    spacer(),

    warningBox([
      [bold('⚠️  Never reuse a Charge Box ID ', { size: 20, color: AMBER }),
       run('even if a charger is decommissioned. Append -R2 for reinstalled hardware.', { size: 20 })],
    ]),
    spacer(200),

    // ── Step-by-step ──────────────────────────────────
    h1('3.  Step-by-Step Procedure'),
    spacer(80),

    stepBox(1, 'Register charger in SteVe', [
      para([run('Log into the SteVe admin panel:', { size: 20 })], { spacing: { after: 60 } }),
      codeBlock(['http://136.113.7.146:8080/steve']),
      spacer(80),
      para([run('Navigate to  Charge Points  →  Add New Charge Point', { size: 20 })], { spacing: { after: 60 } }),
      twoColTable([
        tableHeader('SteVe Field', 'Value to enter'),
        twoCol('chargeBoxId',   'Exact ID from §2.2  (e.g. CS-DC50K-PUN-001)'),
        twoCol('description',   'Site name + address'),
        twoCol('locationLat',   'GPS latitude (6 decimal places)'),
        twoCol('locationLong',  'GPS longitude (6 decimal places)'),
        twoCol('note',          'Operator name + contact number'),
      ]),
      spacer(80),
      para([run('Click  Save.  The charger will appear in the charge_box table in stevedb.', { size: 20 })]),
    ]),
    spacer(),

    stepBox(2, 'Power charger on and verify OCPP connection', [
      para([run(
        'The charger must connect to SteVe via OCPP WebSocket before the next steps. ' +
        'Confirm the OCPP endpoint on the charger is set to:', { size: 20 }
      )]),
      spacer(60),
      codeBlock(['ws://136.113.7.146:8080/steve/websocket/CentralSystemService/{chargeBoxId}']),
      spacer(80),
      para([run('Verify connection in SteVe:', { size: 20 })]),
      bullet('Charge Points list shows the new charger'),
      bullet('Status column shows  Available  (not  Unknown  or  Unavailable)'),
      bullet('Last heartbeat timestamp is recent (within 60 seconds)'),
      spacer(80),
      dangerBox([
        [bold('🚫  Do NOT proceed ', { size: 20, color: RED }),
         run('if the charger shows Unknown or has no heartbeat. The OCPP connection must be live.', { size: 20 })],
      ]),
    ]),
    spacer(),

    stepBox(3, 'Add charger_capabilities to app database', [
      para([run(
        'SSH into the GCP server and run the following MySQL commands. ' +
        'Replace placeholder values with actual data for this charger:', { size: 20 }
      )]),
      spacer(80),
      codeBlock([
        'ssh root@136.113.7.146',
        'mysql -u $APP_DB_USER -p"$APP_DB_PASSWORD" voltstartev_db',
        '',
        '-- Add one row per connector',
        '-- connector_id starts at 1, increment for each connector',
        '',
        'INSERT INTO charger_capabilities (',
        '  charge_box_id, connector_id, max_power_watts,',
        '  power_type, connector_type, vehicle_category',
        ') VALUES',
        "  ('CS-DC50K-PUN-001', 1, 50000, 'DC', 'CCS2',  '4W'),",
        "  ('CS-DC50K-PUN-001', 2, 50000, 'DC', 'CHAdeMO','4W');",
        '',
        '-- Verify',
        "SELECT * FROM charger_capabilities WHERE charge_box_id = 'CS-DC50K-PUN-001';",
      ]),
      spacer(80),
      h3('Connector type reference:'),
      table3([
        head3('connector_type value', 'Physical connector', 'Typical vehicles'),
        row3('CCS2',     'CCS Combo 2',       'Nexon EV, ZS EV, Tiago EV, most new 4W'),
        row3('Type2',    'Mennekes / IEC',    'AC charging, most 4-wheelers'),
        row3('CHAdeMO',  'CHAdeMO',           'Older imports, some Kia / Hyundai'),
        row3('GBT',      'GB/T',              'BYD, some Chinese EVs'),
        row3('BharatAC', 'Bharat AC-001',     '3-wheelers, 2-wheelers (15A socket)'),
        row3('BharatDC', 'Bharat DC-001',     '3-wheelers DC fast charge'),
        row3('ThreePin', '3-pin 5A/15A',      '2-wheelers, basic slow charging'),
        row3('Type1',    'SAE J1772',         'Older American/Japanese imports'),
      ]),
    ]),
    spacer(),

    stepBox(4, 'Add charger_config (legacy fallback)', [
      para([run(
        'This table is used as a fallback by older API paths. ' +
        'Add one row per charger (not per connector):', { size: 20 }
      )]),
      spacer(80),
      codeBlock([
        'INSERT INTO charger_config (',
        '  charge_box_id, max_power_w, power_type',
        ') VALUES (',
        "  'CS-DC50K-PUN-001', 50000, 'DC'",
        ')',
        'ON DUPLICATE KEY UPDATE',
        '  max_power_w = VALUES(max_power_w),',
        '  power_type  = VALUES(power_type);',
      ]),
    ]),
    spacer(),

    stepBox(5, 'Add pricing (see Pricing Document for full reference)', [
      para([run(
        'Every charger must have at least one active pricing row. ' +
        'Without pricing the charger will show on the map but sessions cannot be costed. ' +
        'See the companion document  "Add / Update Pricing"  for all pricing models.',
        { size: 20 }
      )]),
      spacer(80),
      para([run('Minimum required — simple per_kwh example:', { size: 20 })]),
      spacer(60),
      codeBlock([
        'INSERT INTO charger_pricing (',
        '  charge_box_id, connector_id, pricing_model,',
        '  rate_per_kwh, session_fee, is_active, valid_from',
        ') VALUES (',
        "  'CS-DC50K-PUN-001', NULL, 'per_kwh',",
        '   14.00, 25.00, 1, NOW()',
        ');',
      ]),
      spacer(80),
      infoBox([
        [bold('connector_id = NULL ', { size: 20 }), run('means the pricing applies to all connectors on this charger. ', { size: 20 })],
        [run('Set a specific connector_id to override pricing for one connector only.', { size: 20 })],
      ]),
    ]),
    spacer(),

    stepBox(6, 'Invalidate server-side caches', [
      para([run(
        'The charger list is cached in memory for 10 minutes (charger_pricing) ' +
        'and 30 minutes (charger_config). Restart the backend to clear immediately:',
        { size: 20 }
      )]),
      spacer(80),
      codeBlock([
        'ssh root@136.113.7.146',
        'cd /build/VoltStartEV_Backend',
        'pm2 restart voltstartev',
        '',
        '# Or if running via npm:',
        'pm2 restart all',
        '',
        '# Verify server is back:',
        'curl http://136.113.7.146:3000/',
      ]),
    ]),
    spacer(),

    stepBox(7, 'Verify end-to-end in mobile app', [
      bullet('Open VoltStartEV app'),
      bullet('Pull down on map to refresh charger list'),
      bullet('New charger marker appears on map at correct location'),
      bullet('Tap marker — modal shows correct connector types and pricing'),
      bullet('Select connector — pricing estimate calculates correctly'),
      bullet('Start test session with qatest001 / QATEST001'),
      bullet('Session screen shows live telemetry within 15 seconds'),
      bullet('Stop session — final cost appears in History screen'),
      bullet('Wallet balance reduced by correct amount'),
    ]),
    spacer(),

    successBox([
      [bold('✅  Charger is live ', { size: 20, color: GREEN }),
       run('when all 8 checks in Step 7 pass. Log the charger ID and date in the network register.', { size: 20 })],
    ]),
    spacer(200),

    // ── Troubleshooting ───────────────────────────────
    h1('4.  Troubleshooting'),
    spacer(80),

    table3([
      head3('Problem', 'Likely cause', 'Fix'),
      row3(
        'Charger not appearing on map',
        'Missing from charger_capabilities or cache not cleared',
        'Check Step 3 then pm2 restart'
      ),
      row3(
        'Status shows Unknown',
        'OCPP connection not established',
        'Check charger OCPP endpoint URL matches Step 2'
      ),
      row3(
        'Pricing shows Contact operator',
        'No active row in charger_pricing',
        'Complete Step 5 — at least one per_kwh row required'
      ),
      row3(
        'Start session returns Failed',
        'Charger not connected or tag not registered',
        'Check SteVe heartbeat; verify OCPP tag exists in stevedb.ocpp_tag'
      ),
      row3(
        'costSoFar shows 0 in session',
        'rate_per_kwh is NULL in charging_sessions',
        'Pricing INSERT failed — re-run Step 5 and restart backend'
      ),
      row3(
        'Wrong power shown in filter',
        'max_power_watts wrong in charger_capabilities',
        'UPDATE charger_capabilities SET max_power_watts = ? WHERE charge_box_id = ?'
      ),
    ]),
    spacer(200),

    // ── Rollback ──────────────────────────────────────
    h1('5.  Rollback Procedure'),
    para([run(
      'If the charger must be removed from the network (decommissioned, faulty, or incorrectly added):',
      { size: 20 }
    )]),
    spacer(80),

    warningBox([
      [bold('⚠️  Never hard-delete rows ', { size: 20, color: AMBER }),
       run('from charger_pricing or charging_sessions — this breaks historical records.', { size: 20 })],
    ]),
    spacer(80),

    codeBlock([
      '-- 1. Deactivate pricing (charger disappears from priced list)',
      "UPDATE charger_pricing SET is_active = 0 WHERE charge_box_id = 'CS-DC50K-PUN-001';",
      '',
      '-- 2. Remove capabilities (charger disappears from map)',
      "DELETE FROM charger_capabilities WHERE charge_box_id = 'CS-DC50K-PUN-001';",
      '',
      '-- 3. Remove config',
      "DELETE FROM charger_config WHERE charge_box_id = 'CS-DC50K-PUN-001';",
      '',
      '-- 4. Decommission in SteVe (do NOT delete)',
      '--    Navigate to SteVe admin → Charge Points → Edit → set note = "DECOMMISSIONED YYYY-MM-DD"',
      '',
      '-- 5. Restart backend',
      'pm2 restart voltstartev',
    ]),
    spacer(200),

    // ── Change log ────────────────────────────────────
    h1('6.  Change Log'),
    twoColTable([
      tableHeader('Version / Date', 'Change'),
      twoCol('v1.0 — May 2026',   'Initial document. Covers OCPP registration, capabilities, pricing, verification.'),
    ]),
  ];

  return new Document({
    numbering: {
      config: [
        { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ],
    },
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: NAVY },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: SLATE },
          paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 22, bold: true, font: 'Arial', color: DGRAY },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: IN(1), right: IN(1), bottom: IN(1), left: IN(1) },
        },
      },
      headers: { default: makeHeader('Adding a New Charger') },
      footers: { default: makeFooter() },
      children,
    }],
  });
}

// ════════════════════════════════════════════════════════════════
// DOCUMENT 2 — Add / Update Pricing
// ════════════════════════════════════════════════════════════════
function buildPricingDoc() {
  const children = [

    // ── Cover ─────────────────────────────────────────
    new Paragraph({
      children: [bold('VoltStartEV', { size: 72, color: TEAL })],
      alignment: AlignmentType.CENTER,
      spacing: { before: IN(1), after: 80 },
    }),
    new Paragraph({
      children: [bold('Operations Manual', { size: 36, color: NAVY })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
    }),
    new Paragraph({
      children: [bold('Add / Update Charger Pricing', { size: 44, color: SLATE })],
      alignment: AlignmentType.CENTER,
      border: {
        top:    { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 8 },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 8 },
      },
      spacing: { before: 200, after: 200 },
    }),
    new Paragraph({
      children: [run('Version 1.0   |   May 2026', { size: 20, color: DGRAY })],
      alignment: AlignmentType.CENTER,
      spacing: { after: IN(1) },
    }),
    new Paragraph({ children: [new PageBreak()] }),

    // ── Overview ──────────────────────────────────────
    h1('1.  Overview'),
    para([run(
      'Pricing in VoltStartEV is stored per charger in the charger_pricing table. ' +
      'The backend pricing engine reads this table and applies the correct model when a session ends. ' +
      'The table is cached for 10 minutes — changes take effect after a backend restart or cache expiry.',
      { size: 20 }
    )]),
    spacer(),

    infoBox([
      [bold('Who can do this: ', { size: 20 }), run('VoltStartEV backend administrator only', { size: 20 })],
      [bold('Requires: ', { size: 20 }), run('SSH + MySQL access to GCP server', { size: 20 })],
      [bold('Takes effect: ', { size: 20 }), run('Immediately on next session start (after cache clears or restart)', { size: 20 })],
      [bold('Impact: ', { size: 20 }), run('Ongoing sessions NOT affected. Only new sessions pick up new pricing.', { size: 20 })],
    ]),
    spacer(200),

    // ── Table structure ───────────────────────────────
    h1('2.  Table Structure — charger_pricing'),
    spacer(80),
    twoColTable([
      tableHeader('Column', 'Description'),
      twoCol('charge_box_id',  'Charger ID (e.g. CS-DC50K-PUN-001). Must match SteVe charge_box_id exactly.'),
      twoCol('connector_id',   'NULL = applies to all connectors. Integer = specific connector only.'),
      twoCol('pricing_model',  'One of: per_kwh | per_minute | tiered_power | time_of_use | flat_session'),
      twoCol('rate_per_kwh',   'Price per kWh in INR. Used by per_kwh and tiered_power models.'),
      twoCol('rate_per_minute','Price per minute in INR. Used by per_minute model only.'),
      twoCol('session_fee',    'Fixed amount charged at session start (₹ INR). Default 0.'),
      twoCol('tiers',          'JSON array of power tiers for tiered_power model. NULL for others.'),
      twoCol('tou_config',     'JSON object with peak/off-peak hours for time_of_use. NULL for others.'),
      twoCol('is_active',      '1 = live pricing. 0 = deactivated. Only one active row per charger.'),
      twoCol('valid_from',     'Datetime from which this pricing is effective. Defaults to NOW().'),
      twoCol('valid_until',    'Datetime until pricing expires. NULL = never expires.'),
    ]),
    spacer(200),

    // ── Pricing models ────────────────────────────────
    h1('3.  Pricing Models — Reference'),
    spacer(80),

    // 3.1 per_kwh
    labelStrip('Model 1: per_kwh — Simple flat rate per kilowatt-hour'),
    para([run(
      'The most common model. User pays a fixed rate per kWh of energy delivered, ' +
      'plus an optional session fee. Best for AC chargers and locations with predictable demand.',
      { size: 20 }
    )]),
    spacer(60),
    twoColTable([
      tableHeader('Parameter', 'Value'),
      twoCol('pricing_model',  'per_kwh'),
      twoCol('rate_per_kwh',   'Set to your rate in ₹ (e.g. 8.00)'),
      twoCol('session_fee',    'Optional fixed charge in ₹ (e.g. 0.00 or 5.00)'),
      twoCol('tiers',          'NULL'),
      twoCol('tou_config',     'NULL'),
    ]),
    spacer(80),
    codeBlock([
      '-- AC 22kW charger at ₹9.00/kWh with no session fee',
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  rate_per_kwh, session_fee, is_active, valid_from',
      ') VALUES (',
      "  'CS-AC22K-PUN-001', NULL, 'per_kwh',",
      '   9.00, 0.00, 1, NOW()',
      ');',
    ]),
    spacer(160),

    // 3.2 per_minute
    labelStrip('Model 2: per_minute — Rate per minute of session time', SLATE),
    para([run(
      'Useful for destination chargers where you want to incentivise quick turnaround. ' +
      'Session fee is typically zero. Rate is charged for every minute the session is active.',
      { size: 20 }
    )]),
    spacer(60),
    twoColTable([
      tableHeader('Parameter', 'Value'),
      twoCol('pricing_model',   'per_minute'),
      twoCol('rate_per_minute', 'Set to your rate in ₹ (e.g. 2.00)'),
      twoCol('session_fee',     'Optional (e.g. 0.00)'),
      twoCol('rate_per_kwh',    'NULL (not used)'),
      twoCol('tiers',           'NULL'),
    ]),
    spacer(80),
    codeBlock([
      '-- Destination charger at ₹2.00/minute',
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  rate_per_minute, session_fee, is_active, valid_from',
      ') VALUES (',
      "  'CS-AC22K-DEST-001', NULL, 'per_minute',",
      '   2.00, 0.00, 1, NOW()',
      ');',
    ]),
    spacer(160),

    // 3.3 tiered_power
    labelStrip('Model 3: tiered_power — Rate changes based on power delivered', NAVY),
    para([run(
      'The most accurate model for DC fast chargers and HPC where power varies during the session. ' +
      'Each meter value interval is priced at the rate matching the instantaneous power delivered in that interval. ' +
      'This mirrors real electricity cost from the grid.',
      { size: 20 }
    )]),
    spacer(60),
    twoColTable([
      tableHeader('Parameter', 'Value'),
      twoCol('pricing_model',  'tiered_power'),
      twoCol('tiers',          'JSON array — see format below'),
      twoCol('session_fee',    'Fixed charge in ₹ (e.g. 100.00 for HPC)'),
      twoCol('rate_per_kwh',   'Fallback rate if power exceeds all tiers'),
      twoCol('tou_config',     'NULL'),
    ]),
    spacer(80),
    para([bold('Tiers JSON format:', { size: 20 })]),
    spacer(40),
    codeBlock([
      '-- tiers is a JSON array where max_kw is the UPPER limit for that tier rate',
      '-- Power is compared against each tier in order; first matching tier wins',
      '',
      '[',
      '  { "max_kw": 50,  "rate_per_kwh": 16 },   -- up to 50 kW  → ₹16/kWh',
      '  { "max_kw": 150, "rate_per_kwh": 22 },   -- up to 150 kW → ₹22/kWh',
      '  { "max_kw": 350, "rate_per_kwh": 28 }    -- up to 350 kW → ₹28/kWh',
      ']',
    ]),
    spacer(80),
    para([bold('Full INSERT example — HPC 350kW charger:', { size: 20 })]),
    spacer(40),
    codeBlock([
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  rate_per_kwh, session_fee, tiers, is_active, valid_from',
      ') VALUES (',
      "  'CS-HPC350K-MUM-001', NULL, 'tiered_power',",
      '   28.00, 100.00,',
      "  '[{\"max_kw\":50,\"rate_per_kwh\":16},{\"max_kw\":150,\"rate_per_kwh\":22},{\"max_kw\":350,\"rate_per_kwh\":28}]',",
      '   1, NOW()',
      ');',
    ]),
    spacer(160),

    // 3.4 time_of_use
    labelStrip('Model 4: time_of_use — Peak and off-peak rates', '7C3AED', WHITE),
    para([run(
      'Two rates based on time of day. Off-peak rate applies during specified hours (e.g. overnight). ' +
      'Peak rate applies at all other times. Useful for grid-friendly charging incentives.',
      { size: 20 }
    )]),
    spacer(60),
    twoColTable([
      tableHeader('Parameter', 'Value'),
      twoCol('pricing_model',  'time_of_use'),
      twoCol('tou_config',     'JSON object — see format below'),
      twoCol('session_fee',    'Optional fixed charge'),
      twoCol('rate_per_kwh',   'Fallback if tou_config is NULL'),
      twoCol('tiers',          'NULL'),
    ]),
    spacer(80),
    codeBlock([
      '-- tou_config format:',
      '{',
      '  "peak_rate":     14.00,   -- ₹/kWh during peak hours',
      '  "offpeak_rate":   6.00,   -- ₹/kWh during off-peak hours',
      '  "offpeak_start": "22:00", -- Off-peak starts (24hr format, IST)',
      '  "offpeak_end":   "06:00"  -- Off-peak ends',
      '}',
      '',
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  rate_per_kwh, session_fee, tou_config, is_active, valid_from',
      ') VALUES (',
      "  'CS-AC22K-SMART-001', NULL, 'time_of_use',",
      '   14.00, 0.00,',
      '  \'{"peak_rate":14.00,"offpeak_rate":6.00,"offpeak_start":"22:00","offpeak_end":"06:00"}\',',
      '   1, NOW()',
      ');',
    ]),
    spacer(160),

    // 3.5 flat_session
    labelStrip('Model 5: flat_session — Fixed price per session regardless of energy', '059669', WHITE),
    para([run(
      'User pays a single flat fee per session. Useful for parking-integrated chargers ' +
      'where the session fee is bundled with parking. Energy is not metered for billing.',
      { size: 20 }
    )]),
    spacer(60),
    codeBlock([
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  session_fee, rate_per_kwh, is_active, valid_from',
      ') VALUES (',
      "  'CS-AC7K-PARK-001', NULL, 'flat_session',",
      '   50.00, 0.00, 1, NOW()',
      ');',
    ]),
    spacer(200),

    // ── Update pricing ────────────────────────────────
    h1('4.  Updating Existing Pricing'),
    para([run(
      'Never UPDATE an active pricing row directly. Instead, deactivate the old row and insert a new one. ' +
      'This preserves the audit trail — historical sessions always reference the rate that was live when they ran.',
      { size: 20 }
    )]),
    spacer(80),

    warningBox([
      [bold('⚠️  Never UPDATE rate_per_kwh on an active row ', { size: 20, color: AMBER }),
       run('— sessions that completed before the update will show wrong costs in the history.', { size: 20 })],
    ]),
    spacer(80),

    h2('4.1  Correct procedure — deactivate old, insert new'),
    codeBlock([
      '-- Step 1: Deactivate current pricing',
      "UPDATE charger_pricing",
      "SET is_active = 0,",
      "    valid_until = NOW()",
      "WHERE charge_box_id = 'CS-AC22K-PUN-001'",
      "  AND is_active = 1;",
      '',
      '-- Step 2: Insert new pricing (takes effect immediately)',
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  rate_per_kwh, session_fee, is_active, valid_from',
      ') VALUES (',
      "  'CS-AC22K-PUN-001', NULL, 'per_kwh',",
      '   11.00, 0.00, 1, NOW()',
      ');',
      '',
      '-- Step 3: Restart backend to clear 10-minute cache',
      '-- ssh root@136.113.7.146 && pm2 restart voltstartev',
    ]),
    spacer(160),

    h2('4.2  Scheduled pricing change (future date)'),
    para([run(
      'To schedule a price change for a future date without downtime:', { size: 20 }
    )]),
    codeBlock([
      '-- Keep old pricing active until the change date',
      "UPDATE charger_pricing",
      "SET valid_until = '2026-06-01 00:00:00'",
      "WHERE charge_box_id = 'CS-AC22K-PUN-001' AND is_active = 1;",
      '',
      '-- Insert new pricing with future valid_from',
      'INSERT INTO charger_pricing (',
      '  charge_box_id, connector_id, pricing_model,',
      '  rate_per_kwh, session_fee, is_active, valid_from',
      ') VALUES (',
      "  'CS-AC22K-PUN-001', NULL, 'per_kwh',",
      "   12.00, 0.00, 1, '2026-06-01 00:00:00'",
      ');',
      '',
      '-- Both rows are in DB. Backend query uses valid_from <= NOW() AND valid_until IS NULL OR valid_until > NOW()',
      '-- New rate activates automatically at midnight on 1 June.',
    ]),
    spacer(200),

    // ── Verification ──────────────────────────────────
    h1('5.  Verification Checklist'),
    spacer(80),

    para([run('After any pricing change, verify the following:', { size: 20 })]),
    spacer(80),

    numbered([bold('Check DB: ', { size: 20 }), run('Confirm only one is_active = 1 row per charger', { size: 20 })]),
    codeBlock([
      "SELECT charge_box_id, pricing_model, rate_per_kwh,",
      "       session_fee, is_active, valid_from, valid_until",
      "FROM charger_pricing",
      "WHERE charge_box_id = 'CS-AC22K-PUN-001'",
      "ORDER BY valid_from DESC;"
    ]),
    spacer(80),

    numbered([bold('Check cache cleared: ', { size: 20 }), run('Restart backend or wait 10 minutes', { size: 20 })]),
    codeBlock(['pm2 restart voltstartev && pm2 logs --lines 20']),
    spacer(80),

    numbered([bold('Check API response: ', { size: 20 }), run('Pricing visible in /api/chargers', { size: 20 })]),
    codeBlock([
      "curl -s http://136.113.7.146:3000/api/chargers | \\",
      "  python3 -m json.tool | grep -A 10 'CS-AC22K-PUN-001'",
    ]),
    spacer(80),

    numbered([bold('Check mobile: ', { size: 20 }), run('Pull to refresh map → tap charger → verify pricing card shows new rate', { size: 20 })]),
    numbered([bold('Check pricing estimate: ', { size: 20 }), run('Select a connector — estimate should use new rate', { size: 20 })]),
    numbered([bold('Check session end: ', { size: 20 }), run('Run a short test session — History shows new rate in cost', { size: 20 })]),
    spacer(200),

    // ── Quick reference ───────────────────────────────
    h1('6.  Quick Reference — Pricing by Charger Type'),
    spacer(80),

    table3([
      head3('Charger type', 'Recommended model', 'Typical rates (India 2026)'),
      row3('AC 3.3–7 kW (2W/3W)', 'per_kwh', '₹6–8/kWh, ₹0 session fee'),
      row3('AC 7–22 kW (4W)',     'per_kwh', '₹8–12/kWh, ₹0–10 session fee'),
      row3('DC 30–50 kW',         'per_kwh or tiered_power', '₹12–18/kWh, ₹20–50 session fee'),
      row3('DC 60–150 kW',        'tiered_power', '₹16–22/kWh, ₹50–100 session fee'),
      row3('HPC 200–350 kW',      'tiered_power', '₹16–28/kWh, ₹100–200 session fee'),
      row3('Destination / hotel', 'per_minute or flat_session', '₹1–3/min or ₹30–80/session'),
      row3('Smart grid / fleet',  'time_of_use', '₹6 off-peak / ₹14 peak'),
    ]),
    spacer(200),

    // ── Change log ────────────────────────────────────
    h1('7.  Change Log'),
    twoColTable([
      tableHeader('Version / Date', 'Change'),
      twoCol('v1.0 — May 2026', 'Initial document. All 5 pricing models documented with full SQL examples.'),
    ]),
  ];

  return new Document({
    numbering: {
      config: [
        { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ],
    },
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: NAVY },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: SLATE },
          paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 22, bold: true, font: 'Arial', color: DGRAY },
          paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: IN(1), right: IN(1), bottom: IN(1), left: IN(1) },
        },
      },
      headers: { default: makeHeader('Add / Update Pricing') },
      footers: { default: makeFooter() },
      children,
    }],
  });
}

// ── Write files ───────────────────────────────────────
async function main() {
  const doc1 = buildAddChargerDoc();
  const buf1 = await Packer.toBuffer(doc1);
  fs.writeFileSync('C:/voltstartEV/VoltStartEV_OPS_Add_Charger.docx', buf1);
  console.log('✅ Written: VoltStartEV_OPS_Add_Charger.docx');

  const doc2 = buildPricingDoc();
  const buf2 = await Packer.toBuffer(doc2);
  fs.writeFileSync('C:/voltstartEV/VoltStartEV_OPS_Pricing_Guide.docx', buf2);
  console.log('✅ Written: VoltStartEV_OPS_Pricing_Guide.docx');
}

main().catch(err => { console.error(err); process.exit(1); });
