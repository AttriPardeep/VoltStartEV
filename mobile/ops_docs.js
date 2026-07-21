const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  TabStopPosition,
} = require('docx');
const fs = require('fs');

const TEAL='0D9488',NAVY='0F172A',SLATE='1E293B',GRAY='64748B',
      LIGHT='F0FDFA',WHITE='FFFFFF',AMBER='F59E0B',RED='EF4444',
      GREEN='10B981',BORDER='CBD5E1';

const bdr=(c=BORDER)=>({style:BorderStyle.SINGLE,size:1,color:c});
const allB=(c)=>({top:bdr(c),bottom:bdr(c),left:bdr(c),right:bdr(c)});

function sp(pt=6){return new Paragraph({children:[new TextRun('')],spacing:{before:pt*20,after:0}});}

function h1(text){return new Paragraph({heading:HeadingLevel.HEADING_1,
  children:[new TextRun({text,font:'Arial',size:36,bold:true,color:NAVY})],
  spacing:{before:480,after:160},
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:TEAL,space:4}}});}

function h2(text){return new Paragraph({heading:HeadingLevel.HEADING_2,
  children:[new TextRun({text,font:'Arial',size:28,bold:true,color:TEAL})],
  spacing:{before:320,after:120}});}

function h3(text){return new Paragraph({heading:HeadingLevel.HEADING_3,
  children:[new TextRun({text,font:'Arial',size:24,bold:true,color:SLATE})],
  spacing:{before:200,after:80}});}

function p(text,opts={}){return new Paragraph({
  children:[new TextRun({text,font:'Arial',size:22,
    color:opts.color||SLATE,bold:opts.bold||false,italics:opts.italic||false})],
  spacing:{before:40,after:40},alignment:opts.align||AlignmentType.LEFT});}

function bullet(text){return new Paragraph({
  numbering:{reference:'bullets',level:0},
  children:[new TextRun({text,font:'Arial',size:22,color:SLATE})],
  spacing:{before:40,after:40}});}

function note(text,type='note'){
  const C={note:{bg:'EFF6FF',b:'3B82F6',lbl:'NOTE',tc:'1D4ED8'},
    warning:{bg:'FFFBEB',b:AMBER,lbl:'WARNING',tc:'92400E'},
    tip:{bg:'F0FDF4',b:GREEN,lbl:'TIP',tc:'166534'},
    danger:{bg:'FEF2F2',b:RED,lbl:'CAUTION',tc:'991B1B'}};
  const c=C[type]||C.note;
  return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
    rows:[new TableRow({children:[new TableCell({
      borders:{top:{style:BorderStyle.SINGLE,size:12,color:c.b},
               bottom:{style:BorderStyle.SINGLE,size:2,color:c.b},
               left:{style:BorderStyle.SINGLE,size:12,color:c.b},
               right:{style:BorderStyle.SINGLE,size:2,color:c.b}},
      shading:{fill:c.bg,type:ShadingType.CLEAR},
      margins:{top:100,bottom:100,left:160,right:160},
      width:{size:9360,type:WidthType.DXA},
      children:[new Paragraph({children:[
        new TextRun({text:c.lbl+'  ',font:'Arial',size:20,bold:true,color:c.tc}),
        new TextRun({text,font:'Arial',size:20,color:SLATE})],
        spacing:{before:0,after:0}})]})]})]});}

function code(lines){return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders:allB('334155'),shading:{fill:'0F172A',type:ShadingType.CLEAR},
    margins:{top:120,bottom:120,left:200,right:200},
    width:{size:9360,type:WidthType.DXA},
    children:lines.map(l=>new Paragraph({
      children:[new TextRun({text:l,font:'Courier New',size:18,color:'22D3EE'})],
      spacing:{before:20,after:20}}))})]})]});}

function tHead(cols,widths){return new TableRow({tableHeader:true,
  children:cols.map((col,i)=>new TableCell({
    borders:allB(TEAL),shading:{fill:TEAL,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:120,right:120},
    width:{size:widths[i],type:WidthType.DXA},
    verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({children:[new TextRun({text:col,font:'Arial',size:20,bold:true,color:WHITE})],
      alignment:AlignmentType.LEFT})]}))});}

function tRow(cells,widths,shade=false){return new TableRow({
  children:cells.map((cell,i)=>{
    const s=String(cell);
    const isReq=s.includes('REQUIRED');
    const isOpt=s.includes('optional');
    return new TableCell({
      borders:allB(BORDER),
      shading:{fill:shade?'F8FAFC':WHITE,type:ShadingType.CLEAR},
      margins:{top:70,bottom:70,left:120,right:120},
      width:{size:widths[i],type:WidthType.DXA},
      verticalAlign:VerticalAlign.TOP,
      children:[new Paragraph({children:[new TextRun({
        text:s,font:i===0?'Courier New':'Arial',
        size:i===0?18:20,
        color:isReq?'991B1B':isOpt?GRAY:SLATE,
        bold:isReq})]})]});})});}

function hdr(title){return new Header({children:[new Paragraph({
  children:[new TextRun({text:'VoltStartEV  ',font:'Arial',size:18,bold:true,color:TEAL}),
    new TextRun({text:`Operations Guide — ${title}`,font:'Arial',size:18,color:GRAY})],
  border:{bottom:{style:BorderStyle.SINGLE,size:4,color:TEAL,space:4}}})]});}

function ftr() {
  return new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: 'CONFIDENTIAL — VoltStartEV Internal Use Only  ',
            font: 'Arial',
            size: 16,
            color: GRAY,
          }),

          new TextRun({
            text: 'Page ',
            font: 'Arial',
            size: 16,
            color: GRAY,
          }),

          PageNumber.CURRENT,
        ],

        border: {
          top: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: TEAL,
            space: 4,
          },
        },

        spacing: {
          before: 60,
        },
      }),
    ],
  });
}

const numbering={config:[{reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,
  text:'•',alignment:AlignmentType.LEFT,
  style:{paragraph:{indent:{left:720,hanging:360}}}}]}]};

const styles={default:{document:{run:{font:'Arial',size:22}}},paragraphStyles:[
  {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,
    run:{size:36,bold:true,font:'Arial'},paragraph:{spacing:{before:480,after:160},outlineLevel:0}},
  {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,
    run:{size:28,bold:true,font:'Arial'},paragraph:{spacing:{before:320,after:120},outlineLevel:1}},
  {id:'Heading3',name:'Heading 3',basedOn:'Normal',next:'Normal',quickFormat:true,
    run:{size:24,bold:true,font:'Arial'},paragraph:{spacing:{before:200,after:80},outlineLevel:2}}]};

const pageProps={size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}};

// ══════════════════════════════════════════════════════════════════
// DOC 1 — ADD CHARGER
// ══════════════════════════════════════════════════════════════════
function buildChargerDoc(){
  const children=[
    new Paragraph({children:[new TextRun({text:'Operations Guide',font:'Arial',size:48,bold:true,color:TEAL})],spacing:{before:0,after:60}}),
    new Paragraph({children:[new TextRun({text:'Adding a New Charger',font:'Arial',size:36,bold:true,color:NAVY})],spacing:{before:0,after:80}}),
    new Paragraph({children:[
      new TextRun({text:'Version 1.0  ',font:'Arial',size:20,color:GRAY}),
      new TextRun({text:'|  May 2026  |  VoltStartEV Engineering',font:'Arial',size:20,color:GRAY})],
      spacing:{before:0,after:0}}),
    new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:8,color:TEAL,space:8}},children:[],spacing:{before:160,after:320}}),

    h1('Overview'),
    p('This document describes the complete procedure for adding a new EV charging station to the VoltStartEV network. Follow every step in order — skipping steps will result in the charger appearing offline or failing to accept charging sessions.'),
    sp(8),
    p('The process updates three systems:'),
    bullet('SteVe OCPP Server — registers the charger for OCPP 1.6 communication'),
    bullet('VoltStartEV Database — adds charger metadata, capabilities, connector types, and pricing'),
    bullet('Mobile App — reflects changes automatically within 10 minutes via cache TTL'),
    sp(8),
    note('All commands must be run on the VoltStartEV backend server (136.113.7.146) as root or the deploy user.','warning'),
    sp(8),

    h1('Prerequisites'),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[480,5040,3840],rows:[
      tHead(['','Requirement','Verify At'],[480,5040,3840]),
      ...([
        ['☐','Charger physically installed and powered on','On-site inspection'],
        ['☐','Network connection active (WiFi or SIM)','Ping the charger IP'],
        ['☐','OCPP URL configured: ws://136.113.7.146:8080/steve/websocket/CentralSystemService','Charger web interface'],
        ['☐','Charge Box ID format decided (see Section 3)','This document'],
        ['☐','Number of connectors and connector types confirmed','Physical inspection'],
        ['☐','Maximum power per connector known (kW)','Charger nameplate / datasheet'],
        ['☐','GPS coordinates ready (6 decimal places)','Google Maps'],
        ['☐','Pricing model decided (see companion Pricing Guide)','Operations team'],
        ['☐','MySQL access to server available','SSH into 136.113.7.146'],
      ]).map((r,i)=>tRow(r,[480,5040,3840],i%2===1))]}),
    sp(16),

    h1('Step 1 — Decide the Charge Box ID'),
    p('Every charger needs a unique Charge Box ID. This ID is used in OCPP, the database, and the mobile app. It cannot be changed without disrupting active sessions.'),
    sp(8),
    h3('Naming Convention'),
    code([
      'Format:  CS-{TYPE}{POWER}-{SEQUENCE}',
      '',
      'Examples:',
      '  CS-AC7K-00003       7 kW AC charger, unit 3',
      '  CS-AC22K-00001      22 kW AC charger, unit 1',
      '  CS-DC50K-00002      50 kW DC charger, unit 2',
      '  CS-DC150K-00001     150 kW DC fast charger, unit 1',
      '  CS-HPC350K-00002    350 kW HPC ultra-fast, unit 2',
      '  CS-BHARAT-00001     Bharat AC/DC standard (2W/3W), unit 1',
      '  CS-GBT60K-00001     60 kW GB/T (BYD/Chinese EVs), unit 1',
    ]),
    sp(8),
    note('Check existing IDs first:\n  SELECT charge_box_id FROM charger_config ORDER BY charge_box_id;','tip'),
    sp(8),
    h3('Type Code Reference'),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1600,1600,2000,4160],rows:[
      tHead(['Code','Max Power','Connector Type','Compatible Vehicles'],[1600,1600,2000,4160]),
      ...([
        ['AC7K','7.2 kW','Type 2 / Bharat AC','Nexon EV, ZS EV, Atto 3, most 4W'],
        ['AC11K','11 kW','Type 2','Tata Punch EV, premium 4W'],
        ['AC22K','22 kW','Type 2','Tesla, Kia EV6, premium 4W'],
        ['DC50K','50 kW','CCS2 + CHAdeMO','All DC-capable 4W'],
        ['DC150K','150 kW','CCS2','Nexon EV Max, Tata Punch EV'],
        ['HPC350K','350 kW','CCS2','Kia EV6, BYD Atto 3, premium'],
        ['BHARAT','15 kW','Bharat AC-001 / DC-001','2-Wheeler, 3-Wheeler, E-Rickshaw'],
        ['GBT','60 kW','GB/T','BYD models, Chinese imports'],
      ]).map((r,i)=>tRow(r,[1600,1600,2000,4160],i%2===1))]}),
    sp(16),

    h1('Step 2 — Register in SteVe OCPP Server'),
    p('SteVe handles the physical OCPP protocol. The charger must exist in SteVe before it can connect.'),
    sp(8),
    note('SteVe web interface: http://136.113.7.146:8080/steve  —  Login: admin / admin','note'),
    sp(8),
    h3('Option A — SteVe Web Interface (recommended)'),
    bullet('Open http://136.113.7.146:8080/steve → Charge Points → Add New'),
    bullet('Enter the Charge Box ID exactly as decided in Step 1'),
    bullet('Set Registration Status = Accepted'),
    bullet('Click Save'),
    sp(8),
    h3('Option B — Direct MySQL'),
    code([
      'mysql -u root -p stevedb -e "',
      'INSERT INTO charge_box (',
      '  charge_box_id,',
      '  registration_status,',
      '  insert_connector_status_after_transaction_msg',
      ') VALUES (',
      "  'CS-AC22K-00003',",
      "  'Accepted',",
      '  0',
      ');',
      '"',
    ]),
    sp(8),
    note('Physically restart the charger after registration so it sends a BootNotification to SteVe. The charger will not appear online until SteVe receives BootNotification.','warning'),
    sp(16),

    h1('Step 3 — Add to App Database'),
    p('The app database has two tables to update: charger_config (metadata + location) and charger_capabilities (per-connector detail including connector type and vehicle compatibility).'),
    sp(8),
    note('charger_capabilities does not exist yet in the database. Run the CREATE TABLE in Step 3a before inserting data.','danger'),
    sp(8),

    h3('3a — Create charger_capabilities table (one-time setup)'),
    code([
      'mysql -u $APP_DB_USER -p"$APP_DB_PASSWORD" $APP_DB_NAME -e "',
      'CREATE TABLE IF NOT EXISTS charger_capabilities (',
      '  id               INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,',
      '  charge_box_id    VARCHAR(255) NOT NULL,',
      '  connector_id     INT NOT NULL DEFAULT 1,',
      '  max_power_watts  INT NOT NULL,',
      "  power_type       ENUM('AC','DC') NOT NULL DEFAULT 'AC',",
      "  connector_type   ENUM('Type2','CCS2','CHAdeMO','GBT','Type1',",
      "                        'BharatAC','BharatDC','ThreePin') NOT NULL DEFAULT 'Type2',",
      "  vehicle_category SET('2W','3W','4W','Bus','Truck') NOT NULL DEFAULT '4W',",
      '  created_at       DATETIME DEFAULT NOW(),',
      '  PRIMARY KEY (id),',
      '  UNIQUE KEY uq_charger_connector (charge_box_id, connector_id),',
      '  INDEX idx_charge_box (charge_box_id)',
      ');',
      '"',
    ]),
    sp(8),
    note('After creating the table, also add location columns to charger_config (Step 3b).','tip'),
    sp(8),

    h3('3b — Add location columns to charger_config (one-time migration)'),
    code([
      'mysql -u $APP_DB_USER -p"$APP_DB_PASSWORD" $APP_DB_NAME -e "',
      'ALTER TABLE charger_config',
      "  ADD COLUMN IF NOT EXISTS street    VARCHAR(255) NULL,",
      "  ADD COLUMN IF NOT EXISTS city      VARCHAR(100) NULL,",
      "  ADD COLUMN IF NOT EXISTS state     VARCHAR(100) NULL,",
      "  ADD COLUMN IF NOT EXISTS pincode   VARCHAR(10)  NULL,",
      "  ADD COLUMN IF NOT EXISTS latitude  DECIMAL(10,6) NULL,",
      "  ADD COLUMN IF NOT EXISTS longitude DECIMAL(10,6) NULL,",
      "  ADD COLUMN IF NOT EXISTS is_public TINYINT(1) DEFAULT 1,",
      "  ADD COLUMN IF NOT EXISTS is_active TINYINT(1) DEFAULT 1;",
      '"',
    ]),
    sp(8),

    h3('3c — Insert charger_config row'),
    code([
      'mysql -u $APP_DB_USER -p"$APP_DB_PASSWORD" $APP_DB_NAME -e "',
      'INSERT INTO charger_config (',
      '  charge_box_id,',
      '  display_name,',
      '  max_power_w,',
      '  power_type,',
      '  street,',
      '  city,',
      '  state,',
      '  pincode,',
      '  latitude,',
      '  longitude,',
      '  is_public,',
      '  is_active',
      ') VALUES (',
      "  'CS-AC22K-00003',            -- Must match SteVe charge_box_id exactly",
      "  'Koregaon Park Charger 3',   -- Display name shown in app",
      '   22000,                      -- Max power in Watts (22 kW = 22000)',
      "  'AC_3_PHASE',                -- AC_3_PHASE | AC_1_PHASE | DC",
      "  '12, North Main Road',       -- Street address",
      "  'Pune',",
      "  'Maharashtra',",
      "  '411001',",
      '   18.537600,                  -- Latitude — 6 decimal places from Google Maps',
      '   73.893800,                  -- Longitude — 6 decimal places from Google Maps',
      '   1,                          -- 1 = public, 0 = private/restricted',
      '   1                           -- 1 = visible in app, 0 = hidden',
      ');',
      '"',
    ]),
    sp(8),

    h3('3d — Insert charger_capabilities rows (one per connector)'),
    p('Add one row per physical connector socket. A dual-port charger needs two INSERT statements with connector_id 1 and 2.'),
    sp(6),
    code([
      'mysql -u $APP_DB_USER -p"$APP_DB_PASSWORD" $APP_DB_NAME -e "',
      'INSERT INTO charger_capabilities',
      '  (charge_box_id, connector_id, max_power_watts, power_type, connector_type, vehicle_category)',
      'VALUES',
      "  ('CS-AC22K-00003', 1, 22000, 'AC', 'Type2', '4W'),",
      "  ('CS-AC22K-00003', 2, 22000, 'AC', 'Type2', '4W');",
      '"',
      '',
      '-- For a Bharat AC/DC charger (2W + 3W compatible):',
      'INSERT INTO charger_capabilities',
      '  (charge_box_id, connector_id, max_power_watts, power_type, connector_type, vehicle_category)',
      'VALUES',
      "  ('CS-BHARAT-00001', 1, 7200,  'AC', 'BharatAC', '2W,3W'),",
      "  ('CS-BHARAT-00001', 2, 15000, 'DC', 'BharatDC', '3W');",
      '"',
    ]),
    sp(8),

    h3('Connector Type Quick Reference'),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[1800,2200,1400,3960],rows:[
      tHead(['connector_type','Full Standard Name','power_type','vehicle_category values'],[1800,2200,1400,3960]),
      ...([
        ['Type2','IEC 62196-2 (Mennekes)','AC','4W'],
        ['CCS2','Combined Charging System 2','DC','4W'],
        ['CHAdeMO','CHAdeMO DC standard','DC','4W'],
        ['GBT','GB/T 20234 (Chinese standard)','DC','4W'],
        ['Type1','SAE J1772 Type 1','AC','4W'],
        ['BharatAC','IS 17017-2-3 Bharat AC-001','AC','2W,3W'],
        ['BharatDC','IS 17017-2-3 Bharat DC-001','DC','2W,3W'],
        ['ThreePin','3-Pin 5A/15A socket','AC','2W,3W'],
      ]).map((r,i)=>tRow(r,[1800,2200,1400,3960],i%2===1))]}),
    sp(16),

    h1('Step 4 — Add Pricing'),
    p('Every charger must have at least one active pricing record. Without pricing, the app shows "Contact operator for pricing details" and the cost estimate will not work.'),
    sp(8),
    note('See the companion document VoltStartEV — Pricing Configuration Guide for all five pricing models. The quickest setup is shown below.','tip'),
    sp(8),
    code([
      '-- Flat per-kWh pricing (simplest setup):',
      'mysql -u $APP_DB_USER -p"$APP_DB_PASSWORD" $APP_DB_NAME -e "',
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_kwh, session_fee, display_name, valid_from, is_active)',
      'VALUES',
      "  ('CS-AC22K-00003', 'per_kwh', 9.00, 0.00, '₹9.00/kWh', NOW(), 1);",
      '"',
    ]),
    sp(16),

    h1('Step 5 — Invalidate Cache and Verify'),
    h3('5a — Restart backend to clear cache'),
    code([
      'cd /build/VoltStartEV_Backend',
      'pm2 restart voltstartev',
      '',
      '# Verify the charger appears in the API:',
      "curl -s http://136.113.7.146:3000/api/chargers | python3 -c \\",
      "  \"import sys,json; [print(c['chargeBoxId']) for c in json.load(sys.stdin)['data']]\" \\",
      '  | grep CS-AC22K-00003',
    ]),
    sp(8),
    h3('5b — Verify OCPP connection'),
    code([
      'mysql -u root -p stevedb -e "',
      'SELECT charge_box_id, last_heartbeat_timestamp,',
      '       registration_status, ocpp_protocol',
      'FROM charge_box',
      "WHERE charge_box_id = 'CS-AC22K-00003';",
      '"',
      '',
      '-- Expected: last_heartbeat within 2 minutes, status = Accepted',
      '',
      'mysql -u root -p stevedb -e "',
      'SELECT c.charge_box_id, c.connector_id, cs.status',
      'FROM connector c',
      'JOIN connector_status cs ON cs.connector_pk = c.connector_pk',
      "WHERE c.charge_box_id = 'CS-AC22K-00003';",
      '"',
      '',
      '-- Expected: each connector shows status = Available',
    ]),
    sp(16),

    h1('Step 6 — End-to-End Test Checklist'),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[480,5520,1800,1560],rows:[
      tHead(['','Test','Expected','Result'],[480,5520,1800,1560]),
      ...([
        ['☐','Charger appears on map in VoltStartEV app','Marker visible',''],
        ['☐','Correct connector type shown in modal','e.g. Type 2',''],
        ['☐','Pricing displayed before starting session','₹X.XX/kWh',''],
        ['☐','Connector status shows Available','Green indicator',''],
        ['☐','Start session via app — session screen appears','kWh updating',''],
        ['☐','Stop session — session completes normally','Summary shown',''],
        ['☐','Session appears in History with correct cost','Cost matches',''],
        ['☐','Wallet balance reduced by correct amount','Balance matches',''],
        ['☐','Connector returns to Available after stop','Green indicator',''],
      ]).map((r,i)=>tRow(r,[480,5520,1800,1560],i%2===1))]}),
    sp(16),

    h1('Troubleshooting'),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[3400,5960],rows:[
      tHead(['Symptom','Resolution'],[3400,5960]),
      ...([
        ['Charger not on app map','Check charger_config.is_active = 1. Restart backend to clear 10-min cache.'],
        ['Connector shows Unavailable','Physically restart charger to trigger BootNotification to SteVe.'],
        ['No pricing shown','Verify charger_pricing row has is_active = 1 and valid_from <= NOW().'],
        ['Session start fails','Check SteVe registration_status = Accepted. Confirm OCPP connection on port 8080.'],
        ['charger_capabilities missing','Run the CREATE TABLE statement in Step 3a.'],
        ['Latitude/longitude columns missing','Run the ALTER TABLE statement in Step 3b.'],
        ['Heartbeat not updating','Check firewall — port 8080 must be open to charger IP.'],
        ['Wrong vehicle types in filter','Update charger_capabilities.vehicle_category for the affected connectors.'],
      ]).map((r,i)=>tRow(r,[3400,5960],i%2===1))]}),
    sp(16),

    h1('Quick Reference — Minimum Required DB Records'),
    code([
      'stevedb.charge_box                   -- 1 row (OCPP registration)',
      'voltstartev_db.charger_config        -- 1 row (location + metadata)',
      'voltstartev_db.charger_capabilities  -- N rows (one per connector)',
      'voltstartev_db.charger_pricing       -- 1+ rows (at least one active)',
    ]),
  ];

  return new Document({numbering,styles,sections:[{
    properties:{page:pageProps},
    headers:{default:hdr('Adding a New Charger')},
    footers:{default:ftr()},
    children}]});
}

// ══════════════════════════════════════════════════════════════════
// DOC 2 — PRICING CONFIGURATION
// ══════════════════════════════════════════════════════════════════
function buildPricingDoc(){
  const children=[
    new Paragraph({children:[new TextRun({text:'Operations Guide',font:'Arial',size:48,bold:true,color:TEAL})],spacing:{before:0,after:60}}),
    new Paragraph({children:[new TextRun({text:'Pricing Configuration Guide',font:'Arial',size:36,bold:true,color:NAVY})],spacing:{before:0,after:80}}),
    new Paragraph({children:[
      new TextRun({text:'Version 1.0  ',font:'Arial',size:20,color:GRAY}),
      new TextRun({text:'|  May 2026  |  VoltStartEV Engineering',font:'Arial',size:20,color:GRAY})],
      spacing:{before:0,after:0}}),
    new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:8,color:TEAL,space:8}},children:[],spacing:{before:160,after:320}}),

    h1('Overview'),
    p('VoltStartEV supports five pricing models. Each charger must have exactly one active pricing record at any time. Pricing is cached in server memory for 10 minutes — changes take up to 10 minutes to appear in the app.'),
    sp(8),
    note('Sessions record their rate_per_kwh and session_fee at the moment the session starts. Changing pricing never affects sessions that are already in progress or completed.','tip'),
    sp(8),
    note('To apply changes immediately without waiting 10 minutes, restart the backend: pm2 restart voltstartev','note'),
    sp(16),

    h1('charger_pricing Table — Column Reference'),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2200,1600,1400,4160],rows:[
      tHead(['Column','Type','Required','Description'],[2200,1600,1400,4160]),
      ...([
        ['charge_box_id','VARCHAR','REQUIRED','Must exactly match SteVe charge_box_id'],
        ['connector_id','INT','optional','NULL = all connectors. Set for per-connector pricing'],
        ['pricing_model','ENUM','REQUIRED','per_kwh | per_minute | per_session | tiered_power | time_of_use | free'],
        ['rate_per_kwh','DECIMAL(8,2)','REQUIRED*','₹ per kWh. Also fallback rate for tiered_power'],
        ['rate_per_minute','DECIMAL(8,2)','optional','Required when pricing_model = per_minute'],
        ['session_fee','DECIMAL(8,2)','optional','Flat fee per session. Defaults to 0.00'],
        ['tiers','JSON','optional','Required when pricing_model = tiered_power'],
        ['tou_config','JSON','optional','Required when pricing_model = time_of_use'],
        ['display_name','VARCHAR(100)','REQUIRED','Shown in app — keep under 20 characters'],
        ['valid_from','DATETIME','REQUIRED','When pricing starts. Use NOW() for immediate'],
        ['valid_until','DATETIME','optional','NULL = never expires. Set when deactivating'],
        ['is_active','TINYINT','REQUIRED','1 = active. Only ONE row per charger should be 1'],
        ['currency','CHAR(3)','optional','Default: INR. Do not change.'],
        ['notes','VARCHAR(255)','optional','Internal notes — not shown in app'],
      ]).map((r,i)=>tRow(r,[2200,1600,1400,4160],i%2===1))]}),
    sp(8),
    p('* rate_per_kwh is required for all models except per_minute and per_session.'),
    sp(16),

    h1('Model 1 — Per kWh  (per_kwh)'),
    p('Simplest model. User pays a fixed amount per kWh delivered. Best for AC chargers where power output is consistent.'),
    sp(6),
    code([
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_kwh, session_fee, display_name, valid_from, is_active)',
      'VALUES',
      "  ('CS-AC7K-00001', 'per_kwh', 7.50, 0.00, '₹7.50/kWh', NOW(), 1);",
      '',
      '-- With session fee:',
      "  ('CS-DC50K-00001', 'per_kwh', 14.00, 25.00, '₹14/kWh + ₹25', NOW(), 1);",
    ]),
    sp(16),

    h1('Model 2 — Per Minute  (per_minute)'),
    p('User pays per minute of connection time. Useful for destination chargers or when preventing long parking occupancy matters more than energy billing.'),
    sp(6),
    code([
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_minute, session_fee, display_name, valid_from, is_active)',
      'VALUES',
      "  ('CS-AC22K-00002', 'per_minute', 2.00, 0.00, '₹2/min', NOW(), 1);",
    ]),
    sp(16),

    h1('Model 3 — Tiered by Power  (tiered_power)'),
    p('Rate changes based on instantaneous power delivered in each 5-second meter interval. Higher power = higher rate per kWh. Used for HPC and DC fast chargers where power varies significantly during a session.'),
    sp(6),
    note('Tiers must be ordered from lowest max_kw to highest. The last tier catches all power above the previous tier\'s threshold.','warning'),
    sp(6),
    code([
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_kwh, session_fee, tiers, display_name, valid_from, is_active)',
      'VALUES (',
      "  'CS-HPC350K-00001',",
      "  'tiered_power',",
      '   16.00,          -- fallback rate if power does not match any tier',
      '   100.00,         -- ₹100 session fee charged once per session',
      "  '[",
      '    {\"max_kw\": 50,  \"rate_per_kwh\": 16},   -- 0-50 kW  → ₹16/kWh',
      '    {\"max_kw\": 150, \"rate_per_kwh\": 22},   -- 50-150 kW → ₹22/kWh',
      '    {\"max_kw\": 350, \"rate_per_kwh\": 28}    -- 150-350 kW → ₹28/kWh',
      "  ]',",
      "  'Tiered ₹16-28/kWh + ₹100',",
      '   NOW(), 1',
      ');',
    ]),
    sp(8),
    p('How tiers work: each meter value interval reports instantaneous power. If that interval reports 80 kW, the ₹22/kWh rate applies to the energy delivered in that interval. The final bill is the sum of all intervals.'),
    sp(16),

    h1('Model 4 — Time of Use  (time_of_use)'),
    p('Different rates at different hours. Peak hours cost more; night charging is discounted. Encourages users to charge during off-peak grid demand periods.'),
    sp(6),
    code([
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_kwh, session_fee, tou_config, display_name, valid_from, is_active)',
      'VALUES (',
      "  'CS-DC50K-00002',",
      "  'time_of_use',",
      '   10.00,    -- default rate (used if hour does not match any window)',
      '   25.00,',
      "  '{",
      '    \"peak\":     {\"hours\": [9,10,11,12,13,14,15,16,17,18], \"rate\": 14.00},',
      '    \"offpeak\":  {\"hours\": [22,23,0,1,2,3,4,5,6],          \"rate\": 6.00},',
      '    \"shoulder\": {\"hours\": [7,8,19,20,21],                  \"rate\": 10.00}',
      "  }',",
      "  'Peak ₹14 / Night ₹6 + ₹25',",
      '   NOW(), 1',
      ');',
    ]),
    sp(16),

    h1('Model 5 — Free  (free)'),
    p('No charge to the user. Used for complimentary charging at hotels, offices, or as a promotional offer. Sessions are still recorded for monitoring purposes.'),
    sp(6),
    code([
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_kwh, session_fee, display_name, valid_from, is_active)',
      'VALUES',
      "  ('CS-AC7K-00001', 'free', 0.00, 0.00, 'FREE', NOW(), 1);",
    ]),
    sp(16),

    h1('Updating Pricing (Correct Procedure)'),
    p('Never UPDATE an existing active pricing row. Always deactivate the old row and INSERT a new one. This preserves the full pricing history.'),
    sp(8),
    code([
      '-- STEP 1: Deactivate current pricing',
      'UPDATE charger_pricing',
      '  SET is_active = 0,',
      '      valid_until = NOW()',
      "WHERE charge_box_id = 'CS-AC7K-00001'",
      '  AND is_active = 1;',
      '',
      '-- STEP 2: Insert new pricing',
      'INSERT INTO charger_pricing',
      '  (charge_box_id, pricing_model, rate_per_kwh, session_fee, display_name, valid_from, is_active)',
      'VALUES',
      "  ('CS-AC7K-00001', 'per_kwh', 8.50, 0.00, '₹8.50/kWh', NOW(), 1);",
      '',
      '-- STEP 3: Clear cache (optional — auto-clears after 10 min)',
      'pm2 restart voltstartev',
    ]),
    sp(8),
    note('If a session is active when you change pricing, the active session continues at the OLD rate. The new rate applies only to sessions started after the change.','warning'),
    sp(16),

    h1('Current Live Pricing — May 2026'),
    p('Update this table whenever pricing changes. Keep as the single source of truth.'),
    sp(8),
    new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[2600,1600,1200,1200,2760],rows:[
      tHead(['Charger ID','Model','Rate/kWh','Sess. Fee','Display in App'],[2600,1600,1200,1200,2760]),
      ...([
        ['CS-AC7K-00001, 00002','per_kwh','₹7.50','—','₹7.50/kWh'],
        ['CS-AC11K-00001','per_kwh','₹8.00','—','₹8.00/kWh'],
        ['CS-AC22K-00001/2/3','per_kwh','₹9.00','—','₹9.00/kWh'],
        ['CS-DC50K-00001','per_kwh','₹14.00','₹25','₹14/kWh + ₹25'],
        ['CS-DC150K-00001, 00002','per_kwh','₹18.00','₹50','₹18/kWh + ₹50'],
        ['CS-CHAD50K-00001','per_kwh','₹14.00','₹25','₹14/kWh + ₹25'],
        ['CS-HPC350K-00001','tiered_power','₹16-28','₹100','Tiered ₹16-28 + ₹100'],
        ['CS-SCHUKO-00001','per_kwh','₹6.00','—','₹6.00/kWh'],
      ]).map((r,i)=>tRow(r,[2600,1600,1200,1200,2760],i%2===1))]}),
    sp(16),

    h1('Verification Queries'),
    code([
      '-- 1. Check all active pricing',
      'SELECT charge_box_id, pricing_model, rate_per_kwh,',
      '       session_fee, display_name, valid_from',
      'FROM charger_pricing',
      'WHERE is_active = 1',
      '  AND valid_from <= NOW()',
      '  AND (valid_until IS NULL OR valid_until > NOW())',
      'ORDER BY charge_box_id;',
      '',
      '-- 2. Check chargers with NO active pricing (these will show "Contact operator")',
      'SELECT cc.charge_box_id',
      'FROM charger_config cc',
      'LEFT JOIN charger_pricing cp',
      '  ON cp.charge_box_id = cc.charge_box_id AND cp.is_active = 1',
      'WHERE cp.id IS NULL AND cc.is_active = 1;',
      '',
      '-- 3. Pricing history for one charger',
      "SELECT id, pricing_model, rate_per_kwh, session_fee,",
      "       valid_from, valid_until, is_active",
      "FROM charger_pricing",
      "WHERE charge_box_id = 'CS-AC7K-00001'",
      "ORDER BY valid_from DESC;",
    ]),
  ];

  return new Document({numbering,styles,sections:[{
    properties:{page:pageProps},
    headers:{default:hdr('Pricing Configuration')},
    footers:{default:ftr()},
    children}]});
}

async function main(){
  const [b1,b2]=await Promise.all([
    Packer.toBuffer(buildChargerDoc()),
    Packer.toBuffer(buildPricingDoc()),
  ]);
  fs.writeFileSync('C:/voltstartEV/mobile/VoltStartEV_OpsGuide_AddCharger.docx',b1);
  fs.writeFileSync('C:/voltstartEV/mobile/VoltStartEV_OpsGuide_Pricing.docx',b2);
  console.log('Done');
}
main().catch(e=>{console.error(e);process.exit(1);});

