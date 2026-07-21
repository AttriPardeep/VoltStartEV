const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  PageBreak,
} = require('docx');
const fs = require('fs');

// ── Palette ──────────────────────────────────────────
const C = {
  navy:'0F172A', slate:'1E293B', teal:'0D9488', tealDim:'0E7490',
  green:'10B981', amber:'F59E0B', red:'EF4444', purple:'7C3AED',
  blue:'3B82F6', gray:'64748B', sub:'94A3B8', white:'FFFFFF',
  border:'CBD5E1', light:'F0FDFA',
};

// ── Helpers ───────────────────────────────────────────
const b  = (c=C.border)=>({style:BorderStyle.SINGLE,size:1,color:c});
const ab = (c=C.border)=>({top:b(c),bottom:b(c),left:b(c),right:b(c)});
const sp = (pt=8)=>new Paragraph({children:[new TextRun('')],spacing:{before:pt*20,after:0}});
const pb = ()=>new Paragraph({children:[new PageBreak()]});

const h1 = t=>new Paragraph({heading:HeadingLevel.HEADING_1,
  children:[new TextRun({text:t,font:'Arial',size:36,bold:true,color:C.navy})],
  spacing:{before:480,after:160},
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:C.teal,space:4}}});

const h2 = t=>new Paragraph({heading:HeadingLevel.HEADING_2,
  children:[new TextRun({text:t,font:'Arial',size:28,bold:true,color:C.teal})],
  spacing:{before:320,after:120}});

const h3 = t=>new Paragraph({heading:HeadingLevel.HEADING_3,
  children:[new TextRun({text:t,font:'Arial',size:24,bold:true,color:C.slate})],
  spacing:{before:200,after:80}});

const p = (t,opts={})=>new Paragraph({
  children:[new TextRun({text:t,font:'Arial',size:22,
    color:opts.color||C.slate,bold:opts.bold||false,italics:opts.italic||false})],
  spacing:{before:40,after:40}});

const bul = (t,ref='bullets')=>new Paragraph({
  numbering:{reference:ref,level:0},
  children:[new TextRun({text:t,font:'Arial',size:22,color:C.slate})],
  spacing:{before:40,after:40}});

const note=(text,type='note')=>{
  const cfg={
    note:{bg:'EFF6FF',bc:'3B82F6',lbl:'NOTE',tc:'1D4ED8'},
    tip:{bg:'F0FDF4',bc:C.green,lbl:'TIP',tc:'166534'},
    warning:{bg:'FFFBEB',bc:C.amber,lbl:'IMPORTANT',tc:'92400E'},
    danger:{bg:'FEF2F2',bc:C.red,lbl:'CRITICAL',tc:'991B1B'},
  };
  const c=cfg[type]||cfg.note;
  return new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
    rows:[new TableRow({children:[new TableCell({
      borders:{top:{style:BorderStyle.SINGLE,size:10,color:c.bc},
        bottom:b(c.bc),left:{style:BorderStyle.SINGLE,size:10,color:c.bc},right:b(c.bc)},
      shading:{fill:c.bg,type:ShadingType.CLEAR},
      margins:{top:90,bottom:90,left:160,right:160},
      width:{size:9360,type:WidthType.DXA},
      children:[new Paragraph({children:[
        new TextRun({text:c.lbl+'  ',font:'Arial',size:20,bold:true,color:c.tc}),
        new TextRun({text,font:'Arial',size:20,color:C.slate})],
        spacing:{before:0,after:0}})]})]})]});
};

const thead=(cols,widths)=>new TableRow({tableHeader:true,
  children:cols.map((col,i)=>new TableCell({
    borders:ab(C.teal),shading:{fill:C.teal,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:120,right:120},
    width:{size:widths[i],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
    children:[new Paragraph({children:[new TextRun({text:col,font:'Arial',size:20,bold:true,color:C.white})],
      alignment:AlignmentType.LEFT})]}))});

const trow = (cells, widths, shade = false) =>
  new TableRow({
    children: cells.map(
      (cell, i) =>
        new TableCell({
          borders: ab(C.border),
          shading: {
            fill: shade ? 'F8FAFC' : C.white,
            type: ShadingType.CLEAR,
          },
          margins: {
            top: 70,
            bottom: 70,
            left: 120,
            right: 120,
          },
          width: {
            size: widths[i],
            type: WidthType.DXA,
          },
          verticalAlign: VerticalAlign.TOP,
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: String(cell),
                  font: i === 0 ? 'Courier New' : 'Arial',
                  size: i === 0 ? 18 : 20,
                  color: i === 0 ? C.teal : C.slate,
                }),
              ],
              spacing: {
                before: 0,
                after: 0,
              },
            }),
          ],
        })
    ),
  });

const mkTable=(cols,widths,rows)=>new Table({
  width:{size:widths.reduce((a,b)=>a+b,0),type:WidthType.DXA},columnWidths:widths,
  rows:[thead(cols,widths),...rows.map((r,i)=>trow(r,widths,i%2===1))]});

const hdr=()=>new Header({children:[new Paragraph({
  children:[
    new TextRun({text:'VoltStartEV  ',font:'Arial',size:18,bold:true,color:C.teal}),
    new TextRun({text:'System Design Document  v1.0',font:'Arial',size:18,color:C.gray})],
  border:{bottom:{style:BorderStyle.SINGLE,size:4,color:C.teal,space:4}}})]});

const ftr = () =>
  new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: 'CONFIDENTIAL — VoltStartEV Internal  ',
            font: 'Arial',
            size: 16,
            color: C.gray,
          }),
          new TextRun({
            text: 'Page ',
            font: 'Arial',
            size: 16,
            color: C.gray,
          }),
          PageNumber.CURRENT,
        ],
        border: {
          top: {
            style: BorderStyle.SINGLE,
            size: 4,
            color: C.teal,
            space: 4,
          },
        },
        spacing: {
          before: 60,
        },
      }),
    ],
  });

// ── ASCII-art style flow boxes using table ────────────────────────
const flowBox=(lines,bg='0F172A',tc='22D3EE')=>new Table({
  width:{size:9360,type:WidthType.DXA},columnWidths:[9360],
  rows:[new TableRow({children:[new TableCell({
    borders:ab('334155'),shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:120,bottom:120,left:200,right:200},
    width:{size:9360,type:WidthType.DXA},
    children:lines.map(l=>new Paragraph({
      children:[new TextRun({text:l,font:'Courier New',size:18,color:tc})],
      spacing:{before:20,after:20}}))})]})]});

// ═════════════════════════════════════════════════════════════════
// CONTENT
// ═════════════════════════════════════════════════════════════════
const children = [

  // ── COVER ──────────────────────────────────────────────────────
  new Paragraph({children:[new TextRun({text:'VoltStartEV',font:'Arial',size:64,bold:true,color:C.teal})],spacing:{before:0,after:60}}),
  new Paragraph({children:[new TextRun({text:'System Design Document',font:'Arial',size:44,bold:true,color:C.navy})],spacing:{before:0,after:80}}),
  new Paragraph({children:[new TextRun({text:'Frontend · Backend · OCPP · Database · Real-Time Architecture',font:'Arial',size:22,color:C.gray})],spacing:{before:0,after:40}}),
  new Paragraph({children:[new TextRun({text:'Version 1.0  |  June 2026  |  VoltStartEV Engineering',font:'Arial',size:20,color:C.gray})],spacing:{before:0,after:0}}),
  new Paragraph({border:{bottom:{style:BorderStyle.SINGLE,size:8,color:C.teal,space:8}},children:[],spacing:{before:200,after:400}}),

  // document info table
  mkTable(['',''],  [2000,7360],[
    ['Document Owner','VoltStartEV Engineering Team'],
    ['Status','Draft — Internal Review'],
    ['Last Updated','June 2026'],
    ['Audience','Developers · QA · Architects · Investors'],
    ['Repo (Backend)','github.com/AttriPardeep/VoltStartEV_Backend'],
    ['Repo (Frontend)','github.com/AttriPardeep/VoltStartEV'],
    ['Server IP','136.113.7.146 (GCP — Mumbai)'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 1. EXECUTIVE SUMMARY
  // ═══════════════════════════════════════════════════════════════
  h1('1. Executive Summary'),
  p('VoltStartEV is a full-stack EV charging network platform consisting of a React Native mobile app, a Node.js/TypeScript backend hosted on Google Cloud Platform, and integration with the SteVe OCPP 1.6J charging station management server. The system enables EV drivers to discover nearby chargers, start and stop charging sessions, monitor real-time telemetry, manage their wallet, and configure RFID cards — all from a single mobile application.'),
  sp(8),
  p('The platform supports multiple pricing models, fleet management, reservation booking, push notifications, and an AI assistant named Volt powered by the Anthropic Claude API. The architecture is designed for production readiness with WebSocket-based real-time updates, in-memory caching, atomic database operations, and OCPP 1.6J compliance.'),
  sp(8),
  mkTable(['Metric','Value'],[2400,6960],[
    ['Mobile Platform','React Native + Expo SDK 53 (iOS + Android)'],
    ['Backend Runtime','Node.js 20 + TypeScript (ESM)'],
    ['OCPP Server','SteVe (Open Source, OCPP 1.6J)'],
    ['Database','MySQL 8.0 (dual schema: stevedb + voltstartev_db)'],
    ['Hosting','GCP Compute Engine — Mumbai region (asia-south1)'],
    ['Real-Time','Native WebSocket (no Socket.io dependency)'],
    ['Payment Gateway','Razorpay (wallet top-up via WebView)'],
    ['Push Notifications','Expo Push Notification Service'],
    ['AI Assistant','Anthropic Claude API (claude-sonnet-4-20250514)'],
    ['Charger Count','12 chargers across 5 types (AC, DC, HPC, CHAdeMO, 3-pin)'],
    ['OCPP Version','OCPP 1.6J (JSON over WebSocket)'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 2. SYSTEM ARCHITECTURE
  // ═══════════════════════════════════════════════════════════════
  h1('2. System Architecture'),
  h2('2.1 High-Level Architecture'),
  flowBox([
    '┌─────────────────────────────────────────────────────────────────────┐',
    '│                        MOBILE APP (React Native)                    │',
    '│  MapScreen  SessionScreen  HistoryScreen  ProfileScreen  Wallet     │',
    '│  chargerStore  sessionStore  authStore  filterStore  walletStore    │',
    '└────────────────────┬───────────────────────┬────────────────────────┘',
    '                     │ REST API               │ WebSocket (ws://)      ',
    '                     │ HTTP/JSON              │ Real-time events       ',
    '┌────────────────────▼───────────────────────▼────────────────────────┐',
    '│                  BACKEND (Node.js + TypeScript)                     │',
    '│  Express Routes → Controllers → Services → Database                │',
    '│  WebSocket Server (ChargingWebSocketService)                       │',
    '│  Webhook Processor (OCPP events from SteVe)                        │',
    '│  Cron Jobs (Reconciliation, Tag Sync, Steve Sync Queue)            │',
    '└──────┬──────────────────────────────────────┬───────────────────────┘',
    '       │ MySQL (appPool)                       │ MySQL (stevePool)      ',
    '       │ 25 connections                        │ 15 connections         ',
    '┌──────▼──────────┐               ┌────────────▼────────────────────── ┐',
    '│  voltstartev_db  │               │            stevedb                  │',
    '│  users           │               │  charge_box  connector              │',
    '│  charging_sessions│              │  connector_status                   │',
    '│  wallets         │               │  transaction_start/stop             │',
    '│  charger_pricing │               │  ocpp_tag  user_ocpp_tag            │',
    '│  user_vehicles   │               │  charging_profile                   │',
    '│  user_tags       │               └─────────────────────────────────── ┘',
    '│  fleets          │                                                     ',
    '└──────────────────┘               ┌────────────────────────────────────┐',
    '                                   │  SteVe OCPP Server (port 8080)     │',
    '                                   │  Manages charger communication      │',
    '                                   │  OCPP 1.6J (JSON over WebSocket)    │',
    '                                   └──────────────┬─────────────────────┘',
    '                                                  │ OCPP 1.6J WebSocket   ',
    '                                   ┌──────────────▼─────────────────────┐',
    '                                   │  EV Chargers / SAP Simulator        │',
    '                                   │  CS-HPC350K-00001 (350kW DC)        │',
    '                                   │  CS-AC7K-00001 (7kW AC)             │',
    '                                   │  CS-DC150K-00001 (150kW DC)         │',
    '                                   │  + 9 more chargers                  │',
    '                                   └────────────────────────────────────┘',
  ]),
  sp(12),

  h2('2.2 Data Flow — Charging Session'),
  flowBox([
    'USER TAPS "START CHARGING"',
    '         │',
    '         ▼',
    'Mobile: POST /api/charging/start',
    '   ├─ Check wallet balance >= ₹50              (402 if insufficient)',
    '   ├─ validateTagForUser (user_tags table)      (403 if invalid/expired)',
    '   ├─ Check max_active_transaction_count        (409 if tag in use)',
    '   ├─ Fetch connector capability (max_power_w)',
    '   └─ Call SteVe: POST /operations/RemoteStartTransaction',
    '              { chargeBoxIdList, connectorId, idTag, chargingProfilePk }',
    '         │',
    '         ▼ (~20-40 seconds)',
    'SteVe → Charger: OCPP RemoteStartTransaction',
    'Charger responds: StartTransaction.req',
    'SteVe fires webhook → POST /api/webhooks/steve',
    '         │',
    '         ▼',
    'handleTransactionStarted():',
    '   ├─ Resolve userId from idTag (user_tags → stevedb.user_ocpp_tag)',
    '   ├─ Fetch pricing (charger_pricing with 10-min cache)',
    '   ├─ Store rate_per_kwh + session_fee at creation time',
    '   └─ INSERT charging_sessions (status=active)',
    '         │',
    '         ▼',
    'WebSocket: emit session_started to userId',
    'Mobile: SessionScreen appears, fetchActiveSession()',
    '         │',
    '         ▼ (every 5 seconds)',
    'Charger → SteVe: MeterValues',
    'SteVe fires webhook → POST /api/webhooks/steve',
    'handleMeterValues():',
    '   ├─ extractTelemetry (energy, power, current, voltage, SoC, temp)',
    '   ├─ Calculate cost (delta kWh × rate, tiered by power level)',
    '   ├─ UPDATE charging_sessions (last_cost, peak_power_w, peak_temp_c)',
    '   ├─ CHECK #1: SOC >= target? → RemoteStop + push notification',
    '   ├─ CHECK #2: balance within ₹5? → RemoteStop + push notification',
    '   └─ emit telemetry:update via WebSocket → SessionScreen updates',
    '         │',
    '         ▼',
    'USER TAPS "STOP" (or auto-stop fires)',
    'POST /api/charging/stop → RemoteStopTransaction',
    'Charger: StopTransaction.req → SteVe webhook',
    'handleTransactionEnded():',
    '   ├─ Calculate total_cost using stored rate',
    '   ├─ UPDATE charging_sessions (status=completed, end_time, total_cost)',
    '   ├─ Deduct from wallet (atomic)',
    '   └─ emit session_completed via WebSocket',
    'Mobile: HistoryScreen refresh, wallet balance updates',
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 3. BACKEND ARCHITECTURE
  // ═══════════════════════════════════════════════════════════════
  h1('3. Backend Architecture'),
  h2('3.1 Technology Stack'),
  mkTable(['Component','Technology','Version','Purpose'],[2000,2400,1200,3760],[
    ['Runtime','Node.js','20.x','JavaScript runtime'],
    ['Language','TypeScript (ESM)','5.x','Type safety, ES modules'],
    ['HTTP Framework','Express.js','4.x','REST API routing'],
    ['WebSocket','ws (native)','8.x','Real-time client communication'],
    ['ORM / DB','mysql2','3.x','MySQL connection pool'],
    ['Auth','jsonwebtoken','9.x','JWT token signing and verification'],
    ['Password','bcrypt','5.x','Password hashing (12 rounds)'],
    ['Payment','Razorpay Node SDK','2.x','Payment order creation'],
    ['AI','Anthropic SDK','latest','Volt AI assistant (Claude)'],
    ['Scheduler','node-cron','3.x','Reconciliation and sync jobs'],
    ['Logging','winston','3.x','Structured JSON logging'],
    ['Validation','zod','3.x','Runtime type validation'],
    ['Process Manager','PM2','latest','Production process management'],
    ['Hosting','GCP Compute Engine','—','136.113.7.146 (Mumbai)'],
  ]),
  sp(12),

  h2('3.2 Directory Structure'),
  flowBox([
    'src/',
    '├── server.ts                  # Express + WebSocket init, DB health check',
    '├── config/',
    '│   ├── database.ts            # appPool (25 conn) + stevePool (15 conn)',
    '│   └── logger.ts              # Winston structured logging',
    '├── routes/',
    '│   ├── users.routes.ts        # Auth, profile, vehicles, RFID',
    '│   ├── chargers.routes.ts     # Charger list, status, pricing, OCPP ops',
    '│   ├── charging.routes.ts     # Start/stop session, active session',
    '│   ├── wallet.routes.ts       # Balance, transactions, Razorpay webhook',
    '│   ├── reservations.routes.ts # ReserveNow, CancelReservation',
    '│   └── fleet.routes.ts        # Fleet CRUD (FLEET_ENABLED flag)',
    '├── middleware/',
    '│   ├── auth.middleware.ts     # JWT validation (authenticateJwt)',
    '│   └── auth.ts                # Legacy MVP auth (SKIP_OTP mode)',
    '├── services/',
    '│   ├── ocpp/',
    '│   │   ├── steve-adapter.ts   # getAllChargers(), charger status logic',
    '│   │   └── ocpp-message-handler.ts',
    '│   ├── events/',
    '│   │   ├── webhook-event-processor.ts  # Core session/billing logic',
    '│   │   └── telemetry-extractor.ts      # OCPP MeterValues parsing',
    '│   ├── billing/',
    '│   │   └── pricing.service.ts          # 5 pricing models + 10-min cache',
    '│   ├── charging/',
    '│   │   └── charging-profile.service.ts # SetChargingProfile, RemoteStart/Stop',
    '│   ├── wallet/',
    '│   │   └── wallet.service.ts           # Razorpay, deductions, balance',
    '│   ├── auth/',
    '│   │   ├── tag.service.ts              # OCPP tag assignment',
    '│   │   └── otp.service.ts              # OTP generation + atomic verify',
    '│   ├── sync/',
    '│   │   └── steve-sync.service.ts       # Tag deletion queue + retry',
    '│   ├── fleet/',
    '│   │   └── fleet.service.ts',
    '│   └── notifications/',
    '│       └── push.service.ts             # Expo push notifications',
    '├── websocket/',
    '│   └── charging.websocket.ts  # ChargingWebSocketService class',
    '├── cache/',
    '│   └── chargerState.ts        # In-memory connector status cache',
    '├── jobs/',
    '│   ├── reconciliation.job.ts  # Every 10 min — fix orphaned sessions',
    '│   └── tagSync.job.ts         # Every 2 min — process steve_sync_queue',
    '└── types/',
    '    └── ocpp-1.6.ts            # Zod schemas for OCPP message types',
  ]),
  sp(12),

  h2('3.3 API Endpoints'),
  h3('Authentication'),
  mkTable(['Method','Endpoint','Auth','Description'],[800,2800,800,4960],[
    ['POST','/api/users/register','None','Register new user, auto-assign OCPP tag'],
    ['POST','/api/users/login','None','Login, return JWT token'],
    ['GET','/api/users/me','JWT','Get current user profile'],
    ['PUT','/api/users/me','JWT','Update profile'],
    ['POST','/api/users/otp/request','None','Request OTP for email verification'],
    ['POST','/api/users/otp/verify','None','Verify OTP (atomic, timing-safe)'],
  ]),
  sp(8),
  h3('Chargers'),
  mkTable(['Method','Endpoint','Auth','Description'],[800,3200,800,4560],[
    ['GET','/api/chargers','JWT','List all chargers with status, pricing, connectors'],
    ['GET','/api/chargers/:id','JWT','Single charger status summary'],
    ['GET','/api/chargers/:id/pricing-estimate','JWT','Estimate session cost for user vehicle'],
    ['GET','/api/chargers/:id/metrics','JWT','OCPP MeterValues history for connector'],
    ['GET','/api/chargers/:id/connectors/:cId/status','JWT','Live connector status (cache-first)'],
    ['POST','/api/chargers/:id/connectors/:cId/charging-profile','JWT Operator','SetChargingProfile — power limit'],
    ['DELETE','/api/chargers/:id/connectors/:cId/charging-profile','JWT Operator','ClearChargingProfile'],
    ['PUT','/api/chargers/:id/availability','JWT Operator','ChangeAvailability (Operative/Inoperative)'],
  ]),
  sp(8),
  h3('Charging Sessions'),
  mkTable(['Method','Endpoint','Auth','Description'],[800,2800,800,4960],[
    ['POST','/api/charging/start','JWT','Start session — wallet check → validate tag → RemoteStart'],
    ['POST','/api/charging/stop','JWT','Stop session — RemoteStop → StopTransaction webhook'],
    ['GET','/api/charging/session/active','JWT','Get current active session'],
    ['GET','/api/charging/sessions','JWT','Session history (limit=20 default)'],
  ]),
  sp(8),
  h3('Wallet & Payments'),
  mkTable(['Method','Endpoint','Auth','Description'],[800,2800,800,4960],[
    ['GET','/api/wallet','JWT','Get wallet balance and details'],
    ['GET','/api/wallet/transactions','JWT','Transaction history'],
    ['POST','/api/wallet/order','JWT','Create Razorpay payment order'],
    ['POST','/api/wallet/webhook','None (HMAC)','Razorpay webhook — signature verified'],
  ]),
  sp(8),
  h3('Vehicles, RFID, Fleet, Reservations'),
  mkTable(['Method','Endpoint','Auth','Description'],[800,3000,800,4760],[
    ['GET/POST','/api/users/me/vehicles','JWT','List or add vehicles'],
    ['PUT/DELETE','/api/users/me/vehicles/:id','JWT','Update or remove vehicle'],
    ['PUT','/api/users/me/vehicles/:id/primary','JWT','Set primary vehicle'],
    ['PUT','/api/users/me/vehicles/:id/target-soc','JWT','Update SOC target (mid-session aware)'],
    ['GET/POST','/api/users/me/rfid','JWT','List or add RFID cards'],
    ['DELETE','/api/users/me/rfid/:id','JWT','Remove card (soft delete + SteVe sync)'],
    ['PUT','/api/users/me/rfid/:id/primary','JWT','Set primary tag'],
    ['GET/POST','/api/reservations','JWT','List or create reservation'],
    ['DELETE','/api/reservations/:id','JWT','Cancel reservation'],
    ['GET','/api/reservations/active','JWT','Current active reservation'],
    ['GET/POST/PUT/DELETE','/api/fleet/*','JWT fleet_admin','Fleet management (FLEET_ENABLED flag)'],
  ]),
  sp(16),

  h2('3.4 WebSocket Events'),
  mkTable(['Direction','Event','Payload','Trigger'],[1200,2200,3200,2760],[
    ['Server → Client','connected','{ userId, username }','On successful WS authentication'],
    ['Server → Client','session_started','{ transactionId, chargeBoxId, connectorId, startTime, meterStart }','StartTransaction webhook received'],
    ['Server → Client','telemetry:update','{ transactionId, meterWh, energyKwh, costSoFar, powerW, currentA, voltageV, socPercent, temperatureC, vehicle }','Every MeterValues webhook (~5s)'],
    ['Server → Client','session_completed','{ transactionId, energyKwh, totalCost, durationMinutes, stopReason }','StopTransaction webhook received'],
    ['Server → Client','soc_target_reached','{ transactionId, currentSoc, targetSoc, vehicle, message }','SOC auto-stop triggered'],
    ['Server → Client','balance_critical','{ transactionId, currentBalance, costSoFar, message }','Low balance auto-stop triggered'],
    ['Server → Client','charger:status','{ chargeBoxId, connectorId, status }','StatusNotification webhook received'],
    ['Server → Client','target_soc_updated','{ targetSoc, sessionId, message }','User updates SOC target during active session'],
    ['Client → Server','authenticate','{ token }','On WS connect (also via ?token= URL param)'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 4. DATABASE DESIGN
  // ═══════════════════════════════════════════════════════════════
  h1('4. Database Design'),
  h2('4.1 Schema Overview'),
  p('The system uses two MySQL 8.0 schemas. stevedb is owned and managed by the SteVe OCPP server. voltstartev_db is the application database managed by the VoltStartEV backend. The backend maintains separate connection pools for each schema.'),
  sp(8),

  h2('4.2 voltstartev_db — Key Tables'),
  h3('users'),
  mkTable(['Column','Type','Notes'],[2000,2000,5360],[
    ['user_id','INT UNSIGNED PK','Auto-increment primary key'],
    ['username','VARCHAR(50) UNIQUE','Login identifier'],
    ['email','VARCHAR(255) UNIQUE','Used for OTP delivery'],
    ['password_hash','VARCHAR(255)','bcrypt 12 rounds'],
    ['role','ENUM','customer | fleet_admin | operator | super_admin'],
    ['push_token','VARCHAR(500)','Expo push token for notifications'],
    ['push_enabled','TINYINT','1=notifications enabled'],
    ['target_soc_percent','TINYINT','Default SOC target (legacy field)'],
  ]),
  sp(8),
  h3('charging_sessions'),
  mkTable(['Column','Type','Notes'],[2400,2000,4960],[
    ['session_id','INT UNSIGNED PK','Auto-increment'],
    ['app_user_id','INT UNSIGNED FK','→ users.user_id'],
    ['steve_transaction_pk','INT','SteVe transaction ID — primary join key'],
    ['charge_box_id','VARCHAR(255)','Charger identifier'],
    ['connector_id','INT','Physical connector number (>0)'],
    ['id_tag','VARCHAR(50)','OCPP tag used to start session'],
    ['status','ENUM','active | completed | interrupted | pending'],
    ['start_time / end_time','DATETIME','UTC session boundaries'],
    ['start_meter_value','DECIMAL(10,3)','Meter reading at session start (Wh)'],
    ['last_meter_value','DECIMAL(10,3)','Last received meter reading (Wh) — billing fallback'],
    ['last_cost','DECIMAL(10,2)','Running cost — updated every MeterValues'],
    ['total_cost','DECIMAL(10,2)','Final cost — written on StopTransaction'],
    ['energy_kwh','DECIMAL(10,4)','Final energy delivered'],
    ['rate_per_kwh','DECIMAL(8,2)','Stored at session creation — never re-looked up'],
    ['session_fee','DECIMAL(8,2)','Flat fee stored at creation (charged once)'],
    ['pricing_model','ENUM','per_kwh | tiered_power | per_minute | time_of_use | free'],
    ['tiers','JSON','Tiered pricing brackets (for tiered_power model)'],
    ['peak_power_w','INT','Maximum power seen during session'],
    ['peak_temp_c','DECIMAL(5,1)','Maximum connector temperature'],
    ['last_reported_soc','TINYINT','Last SOC reading — used for noise filtering'],
    ['stop_reason','VARCHAR(50)','Remote | Local | SOCTargetReached | LowBalance | etc.'],
    ['stop_requested','TINYINT','Atomic flag — prevents duplicate stops'],
    ['payment_status','ENUM','pending | completed | failed'],
  ]),
  sp(8),
  h3('charger_pricing'),
  mkTable(['Column','Type','Notes'],[2000,2000,5360],[
    ['id','INT UNSIGNED PK','Auto-increment'],
    ['charge_box_id','VARCHAR(255)','Target charger'],
    ['connector_id','INT NULLABLE','NULL = all connectors, set for per-connector pricing'],
    ['pricing_model','ENUM','per_kwh | per_minute | per_session | tiered_power | time_of_use | free'],
    ['rate_per_kwh','DECIMAL(8,2)','Base rate or fallback for tiered model'],
    ['session_fee','DECIMAL(8,2)','Per-session flat fee (default 0.00)'],
    ['tiers','JSON','[{max_kw, rate_per_kwh}] array for tiered model'],
    ['tou_config','JSON','{peak:{hours,rate}, offpeak:{hours,rate}} for time_of_use'],
    ['valid_from / valid_until','DATETIME','Pricing validity window'],
    ['is_active','TINYINT(1)','Only ONE row active per charger at a time'],
    ['display_name','VARCHAR(100)','Human-readable rate string shown in app'],
  ]),
  sp(8),
  note('IMPORTANT: Never UPDATE existing pricing rows. Always deactivate (is_active=0) and INSERT new. Sessions store rate at creation time — historical costs are never affected by pricing changes.','warning'),
  sp(8),
  h3('Other Key Tables'),
  mkTable(['Table','Purpose','Key Columns'],[2400,3000,3960],[
    ['user_vehicles','EV profiles per user','brand, model, battery_kwh, target_soc, is_primary'],
    ['user_tags','OCPP tags per user','ocpp_tag_id, tag_type (system/external_rfid/fleet), is_primary, is_active'],
    ['wallets','User wallet balance','user_id, balance (DECIMAL 10,2)'],
    ['wallet_transactions','Transaction ledger','type (credit/debit/refund), amount, session_id, description'],
    ['payment_orders','Razorpay orders','razorpay_order_id, amount, status (pending/completed/failed)'],
    ['charger_config','Charger metadata + location','charge_box_id, display_name, street, city, lat, lng, max_power_w, is_active'],
    ['charger_capabilities','Per-connector specs','charge_box_id, connector_id, max_power_watts, connector_type (ENUM), vehicle_category (SET)'],
    ['app_reservations','Reservation tracking','user_id, steve_reservation_pk, charge_box_id, connector_id'],
    ['fleets / fleet_members / fleet_vehicles','Fleet management','FLEET_ENABLED flag gates all fleet features'],
    ['otp_verifications','OTP lifecycle','email, otp_hash (SHA-256), expires_at, used_at, attempts (max 3)'],
    ['webhook_events','Idempotency for OCPP webhooks','event_id (unique), processed_at'],
    ['steve_sync_queue','RFID deletion retry queue','action, payload, attempts, next_retry (exponential backoff)'],
  ]),
  sp(12),

  h2('4.3 stevedb — Key Tables Used by Backend'),
  mkTable(['Table','Usage'],[2800,6560],[
    ['charge_box','Charger registration, last_heartbeat_timestamp, registration_status'],
    ['connector','Physical connectors (connector_id > 0 are ports, 0 = charger itself)'],
    ['connector_status','Latest OCPP status per connector (Available/Charging/Faulted/etc.)'],
    ['transaction_start','Session start records with id_tag, connector_pk, start_timestamp'],
    ['transaction_stop','Session stop records with stop_reason, meter_stop'],
    ['ocpp_tag','RFID tags: id_tag, expiry_date, max_active_transaction_count'],
    ['ocpp_tag_activity','Live tag state: active_transaction_count, in_transaction, blocked'],
    ['user_ocpp_tag','Mapping: user_pk → ocpp_tag_pk (for tag ownership validation)'],
    ['charging_profile','Stored charging profiles referenced by chargingProfilePk in OCPP ops'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 5. FRONTEND ARCHITECTURE
  // ═══════════════════════════════════════════════════════════════
  h1('5. Frontend Architecture'),
  h2('5.1 Technology Stack'),
  mkTable(['Component','Library','Version','Purpose'],[2000,2800,1200,3360],[
    ['Framework','React Native + Expo','SDK 53','Cross-platform mobile'],
    ['Navigation','React Navigation','6.x','Tab + Stack navigation'],
    ['State Management','Zustand','4.x','Global store (no Redux boilerplate)'],
    ['HTTP Client','Axios','1.x','REST API calls with interceptors'],
    ['Maps','react-native-maps','1.18','Google Maps integration'],
    ['Icons','lucide-react-native','0.383','SVG icon system (no emoji)'],
    ['SVG Rendering','react-native-svg','15.x','Required by lucide-react-native'],
    ['Notifications','expo-notifications','0.29','Push notification registration'],
    ['Secure Storage','expo-secure-store','13.x','Token storage (production)'],
    ['Async Storage','@react-native-async-storage','1.x','Token storage (dev/Expo Go)'],
    ['Location','expo-location','17.x','GPS for nearby charger distance'],
    ['WebView','react-native-webview','13.x','Razorpay payment flow'],
    ['Payments','Razorpay WebView','—','Expo Go SDK53 incompatible with native SDK'],
  ]),
  sp(12),

  h2('5.2 Directory Structure'),
  flowBox([
    'src/',
    '├── screens/',
    '│   ├── MapScreen.tsx          # Map, charger discovery, filters, modal',
    '│   ├── SessionScreen.tsx      # Live session telemetry display',
    '│   ├── HistoryScreen.tsx      # Session history + live cost during session',
    '│   ├── ProfileScreen.tsx      # User info, vehicles, RFID, fleet',
    '│   └── WalletScreen.tsx       # Balance, transactions, Razorpay WebView',
    '├── components/',
    '│   ├── icons/',
    '│   │   ├── index.ts           # AppIcon factory, IconSize, IconColors',
    '│   │   └── DynamicIcons.tsx   # DynamicBatteryIcon, DynamicPowerIcon, etc.',
    '│   ├── BoltIcon.tsx           # Pure-View bolt (no SVG — for map markers)',
    '│   ├── RFIDSection.tsx        # RFID card management',
    '│   ├── FleetSection.tsx       # Fleet UI (fleet_admin role only)',
    '│   ├── TargetSocControl.tsx   # SOC target slider (always visible)',
    '│   ├── RazorpayWebView.tsx    # Razorpay payment WebView wrapper',
    '│   └── VehicleModal.tsx       # Add/edit vehicle form',
    '├── store/',
    '│   ├── authStore.ts           # User, token, login/logout, updateUser',
    '│   ├── sessionStore.ts        # Active session, telemetry, WebSocket',
    '│   ├── chargerStore.ts        # Charger list, location, updateConnectorStatus',
    '│   └── filterStore.ts         # Filter state (availability, power, price, type)',
    '├── utils/',
    '│   ├── api.ts                 # Axios instance with JWT interceptors (lazy require)',
    '│   └── socket.ts              # Bridge: native WS → socket.on/off API',
    '└── services/',
    '    └── notifications.ts       # Local + push notification registration',
  ]),
  sp(12),

  h2('5.3 State Management — Zustand Stores'),
  mkTable(['Store','State','Key Actions'],[2000,3400,4160],[
    ['authStore','user, token','login(), logout(), loadToken(), updateUser()'],
    ['sessionStore','activeSession, telemetry, ws, wsConnected, reconnectAttempts','connectWebSocket(), fetchActiveSession(), startSession(), stopSession()'],
    ['chargerStore','chargers, userLocation, locationPermission','fetchChargers(), requestLocation(), updateConnectorStatus()'],
    ['filterStore','availability, minPower, maxDistance, maxPrice, connectorType[], vehicleType','setFilters(), resetFilters()'],
  ]),
  sp(8),

  h2('5.4 WebSocket Architecture'),
  p('The mobile uses a native WebSocket (not Socket.io). The sessionStore.ts manages the connection lifecycle. A socket.ts utility bridges the native WS to a socket.on/off API so HistoryScreen and other screens can subscribe to events without importing the full store.'),
  sp(6),
  flowBox([
    'sessionStore.connectWebSocket()',
    '   ├─ Read token from authStore (lazy require to break circular dependency)',
    '   ├─ Connect: ws://136.113.7.146:3000/ws/charging?token={JWT}',
    '   │   (token in URL — backend authenticates on HTTP upgrade)',
    '   ├─ onopen: log "WS authenticated"',
    '   ├─ onmessage: parse → emitSocketEvent() → update store state',
    '   │   ├─ telemetry:update → set({ telemetry, lastTelemetryAt })',
    '   │   ├─ session_started → fetchActiveSession()',
    '   │   ├─ session_completed → set({ activeSession: null })',
    '   │   ├─ charger:status → updateConnectorStatus() in chargerStore',
    '   │   ├─ soc_target_reached → show Alert',
    '   │   └─ balance_critical → show Alert + navigate to Wallet',
    '   └─ onclose: exponential backoff reconnect (NOT on code 4001/4002)',
    '',
    'socket.ts (event bridge):',
    '   emitSocketEvent(event, data) → listeners.get(event).forEach(fn => fn(data))',
    '   socket.on(event, handler) → add to listeners Map',
    '   socket.off(event, handler) → remove from listeners Map',
    '',
    'Used by:',
    '   HistoryScreen: socket.on("telemetry:update") → update liveCost',
    '   HistoryScreen: socket.on("session_started") → fetchHistory()',
    '   MapScreen:     socket.on("charger:status") → updateConnectorStatus()',
    '   SessionScreen: socket.on("target_soc_updated") → update target bar',
    '   SessionScreen: socket.on("balance_critical") → show alert',
  ]),
  sp(12),

  h2('5.5 Icon System'),
  p('All emoji have been removed and replaced with lucide-react-native SVG icons. A factory function creates consistent icon components with brand defaults.'),
  sp(6),
  note('Map markers CANNOT use SVG icons (React Native Fabric rendering constraint on Android). Map markers use BoltIcon.tsx (pure View components) instead.','danger'),
  sp(6),
  mkTable(['Component','Props','Use Case'],[2400,2400,4560],[
    ['AppIcon.Zap','size, color','Charging indicators, App Tag badges'],
    ['AppIcon.Card','size, color','RFID card icons'],
    ['AppIcon.FleetBuilding','size, color','Fleet section headers'],
    ['AppIcon.Car','size, color','Vehicle section headers'],
    ['AppIcon.Wallet / Rupee','size, color','Wallet, cost displays'],
    ['DynamicBatteryIcon','soc, charging, size','Session battery state (color changes by SOC%)'],
    ['DynamicPowerIcon','powerKw, size','Session power level (color changes by kW tier)'],
    ['DynamicCostIcon','cost, budget, size','Wallet low balance indicator'],
    ['DynamicStatusIcon','status, isReserved, size','Charger modal status (NOT in Marker)'],
    ['DynamicConnectivityIcon','connected, lastUpdateSeconds, size','WebSocket freshness indicator'],
    ['IconBadge','icon, label, color, background','RFID badges (Primary, Blocked, App Tag, RFID Card)'],
    ['IconRow','icon, label, color, size, gap','Section headers in Profile screen'],
    ['BoltIcon (pure View)','color, size','Map marker bolt — no SVG dependency'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 6. PRICING ENGINE
  // ═══════════════════════════════════════════════════════════════
  h1('6. Pricing Engine'),
  h2('6.1 Pricing Models'),
  mkTable(['Model','Key','Billing Logic','Example'],[1600,1800,3400,2560],[
    ['Per kWh','per_kwh','cost = energy_kwh × rate_per_kwh + session_fee','₹9/kWh + ₹0 session fee'],
    ['Per Minute','per_minute','cost = duration_minutes × rate_per_minute','₹2/min'],
    ['Tiered Power','tiered_power','Rate changes per 5s interval based on instantaneous kW','₹16/kWh <50kW, ₹22/kWh <150kW, ₹28/kWh ≥150kW'],
    ['Time of Use','time_of_use','Rate changes by hour of day (peak/shoulder/offpeak)','₹14/kWh 9am-6pm, ₹6/kWh 10pm-6am'],
    ['Free','free','No charge, session tracked for monitoring','Complimentary hotel/office charging'],
  ]),
  sp(8),

  h2('6.2 Pricing Architecture Decisions'),
  bul('Rate stored at session creation (rate_per_kwh, session_fee in charging_sessions) — changing pricing never affects in-progress or historical sessions'),
  bul('Pricing cached for 10 minutes (pricingCache with TTL) — reduces DB queries on every charger list fetch'),
  bul('Session fee charged on first meter value interval only (isFirstChunk flag)'),
  bul('Tiered power uses instantaneous powerW per interval — each 5-second window is billed independently at the power-appropriate rate'),
  bul('safeCost = Math.max(prevCost, newCost) — monotonic guard prevents cost from decreasing due to backward meter values or out-of-order webhooks'),
  bul('Auto-stop STOP_THRESHOLD = ₹5 — RemoteStop sent when (balance - costSoFar) < ₹5'),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 7. OCPP 1.6J INTEGRATION
  // ═══════════════════════════════════════════════════════════════
  h1('7. OCPP 1.6J Integration'),
  h2('7.1 Message Flow'),
  mkTable(['OCPP Message','Direction','Handler','Business Logic'],[2400,1600,2400,3160],[
    ['BootNotification','Charger→SteVe','SteVe native','Charger registration, marks online'],
    ['Heartbeat','Charger→SteVe','SteVe native','Updates last_heartbeat_timestamp every 60s'],
    ['StatusNotification','Charger→SteVe→Backend','webhook-event-processor','Updates connector status, broadcasts charger:status via WS'],
    ['StartTransaction','Charger→SteVe→Backend','handleTransactionStarted()','Creates charging_sessions row with stored pricing'],
    ['MeterValues','Charger→SteVe→Backend','handleMeterValues()','Updates cost, emits telemetry, checks auto-stop conditions'],
    ['StopTransaction','Charger→SteVe→Backend','handleTransactionEnded()','Completes session, deducts wallet, emits session_completed'],
    ['RemoteStartTransaction','Backend→SteVe→Charger','charging-profile.service','Initiates session with optional chargingProfilePk'],
    ['RemoteStopTransaction','Backend→SteVe→Charger','charging-profile.service','Stops active session'],
    ['SetChargingProfile','Backend→SteVe→Charger','chargers.routes','Power limiting — must use chargingProfilePk not inline profile'],
    ['ClearChargingProfile','Backend→SteVe→Charger','chargers.routes','Remove power limit'],
    ['ChangeAvailability','Backend→SteVe→Charger','chargers.routes','Mark connector Operative or Inoperative'],
    ['ReserveNow','Backend→SteVe→Charger','reservations.routes','Reserve connector for specific tag'],
    ['CancelReservation','Backend→SteVe→Charger','reservations.routes','Release reservation'],
  ]),
  sp(8),

  h2('7.2 SteVe API Constraints'),
  note('CRITICAL: SteVe REST API does NOT accept inline chargingProfile objects. You must first CREATE the profile in stevedb (POST /api/v1/chargingProfiles → returns chargingProfilePk), then reference that PK in all operations. All operation endpoints use /operations/* prefix and require chargeBoxIdList (array) not chargeBoxId (string).','danger'),
  sp(8),
  mkTable(['Operation','Correct SteVe Endpoint','Required Payload Fields'],[2400,3200,3760],[
    ['RemoteStartTransaction','POST /api/v1/operations/RemoteStartTransaction','chargeBoxIdList[], connectorId, idTag, chargingProfilePk?'],
    ['RemoteStopTransaction','POST /api/v1/operations/RemoteStopTransaction','chargeBoxIdList[], transactionId'],
    ['SetChargingProfile','POST /api/v1/operations/SetChargingProfile','chargeBoxIdList[], connectorId, chargingProfilePk'],
    ['ClearChargingProfile','POST /api/v1/operations/ClearChargingProfile','chargeBoxIdList[], connectorId, chargingProfilePurpose, stackLevel, filterType'],
    ['ChangeAvailability','POST /api/v1/operations/ChangeAvailability','chargeBoxIdList[], connectorId, availType (not "type")'],
    ['Create Profile','POST /api/v1/chargingProfiles','chargingProfileKind, chargingProfilePurpose, stackLevel, chargingSchedule'],
  ]),
  sp(12),

  h2('7.3 MeterValues Telemetry Extraction'),
  mkTable(['OCPP Measurand','Internal Field','Unit','Notes'],[3000,2000,1200,3160],[
    ['Energy.Active.Import.Register','meterWh','Wh','Cumulative meter — primary billing source'],
    ['Energy.Active.Import.Interval','intervalEnergyWh','Wh','Per-interval delta — fallback if Register missing'],
    ['Power.Active.Import','powerW','W','Used for tiered pricing tier selection'],
    ['Power.Offered','powerOfferedW','W','Charger capability reporting'],
    ['Current.Import (total or L1+L2+L3)','currentA','A','Per-phase or total'],
    ['Current.Import (L1/L2/L3)','currentL1/L2/L3','A','Three-phase breakdown'],
    ['Current.Offered','currentOfferedA','A','What charger is offering'],
    ['Voltage (or L1-N)','voltageV','V','Phase or line voltage'],
    ['SoC','socPercent','%','Battery state of charge'],
    ['Temperature','temperatureC','°C','Connector temperature'],
    ['EVBatteryTemperature / BatteryTemperature','batteryTemperatureC','°C','Vendor-specific EV battery temp'],
    ['Frequency','frequencyHz','Hz','Grid frequency monitoring'],
    ['Power.Reactive.Import','powerReactiveW','W','Grid quality monitoring'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 8. SECURITY
  // ═══════════════════════════════════════════════════════════════
  h1('8. Security'),
  mkTable(['Area','Implementation','Status'],[2400,4960,1800],[
    ['Authentication','JWT (HS256, 30d expiry) stored in AsyncStorage (dev) / expo-secure-store (prod)','✅ Implemented'],
    ['Password Storage','bcrypt with 12 salt rounds','✅ Implemented'],
    ['WebSocket Auth','Token validated on HTTP upgrade (URL query param ?token=)','✅ Implemented'],
    ['Payment Verification','Razorpay webhook HMAC-SHA256 signature verification','✅ Implemented'],
    ['OTP Verification','Atomic UPDATE (race condition safe), SHA-256 hash, timing-safe comparison, max 3 attempts','✅ Implemented'],
    ['Tag Validation','Checks user_tags (active), stevedb.ocpp_tag (expiry), max_active_transaction_count','✅ Implemented'],
    ['RFID Deletion Sync','Soft delete + SteVe expiry_date=2000-01-01 + user_ocpp_tag removal + retry queue','✅ Implemented'],
    ['Duplicate Payment','Idempotency check on Razorpay order_id prevents double-crediting','✅ Implemented'],
    ['Auto-stop Guard','stop_requested atomic flag prevents duplicate RemoteStop on same session','✅ Implemented'],
    ['Rate Limiting','Not yet implemented on auth endpoints','❌ Pending'],
    ['HTTPS/TLS','HTTP only (no SSL configured)','❌ Pending'],
    ['Token Storage','AsyncStorage in dev — not encrypted','⚠️ Use expo-secure-store in production'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 9. CACHING STRATEGY
  // ═══════════════════════════════════════════════════════════════
  h1('9. Caching Strategy'),
  mkTable(['Cache','TTL','Scope','Invalidation'],[2400,1200,2400,3360],[
    ['charger_config (config + location)','30 min','Module-level in steve-adapter.ts','invalidateChargerConfigCache() or backend restart'],
    ['charger_pricing','10 min','Module-level in pricing.service.ts','invalidatePricingCache() or backend restart'],
    ['charger_capabilities','30 min','Module-level in steve-adapter.ts','invalidateCapabilityCache() or backend restart'],
    ['Connector status (per charger)','Until updated','chargerState.ts in-memory Map','updateFromOCPP() on every StatusNotification'],
    ['Mobile charger list','2 min polling','chargerStore.chargers array','fetchChargers() interval + charger:status WS events'],
    ['DB connection pools','Persistent','appPool (25), stevePool (15)','Connection recycled after idleTimeout (60s)'],
  ]),
  sp(8),
  note('Cache logs: "CHARGER CONFIG CACHE HIT", "PRICING CACHE HIT", "CAPABILITIES CACHE REFRESHED" visible in backend logs. These confirm correct caching behavior.','tip'),
  sp(16),

  // ═══════════════════════════════════════════════════════════════
  // 10. BACKGROUND JOBS
  // ═══════════════════════════════════════════════════════════════
  h1('10. Background Jobs'),
  mkTable(['Job','Schedule','Purpose','Key Logic'],[2400,1200,2800,3160],[
    ['Reconciliation','Every 10 min','Fix orphaned sessions (backend restart mid-session)','Find active sessions with no MeterValues in 15 min → mark completed at last_cost'],
    ['Tag Sync Queue','Every 2 min','Process failed RFID deletion syncs to SteVe','Exponential backoff: 2min→4min→8min→16min→32min (max 5 attempts)'],
    ['Steve Sync Queue','Every 2 min','Process any queued SteVe operations','Retry on SteVe downtime'],
    ['Concurrency Guard','Per job run','Prevent job overlap','isRunning flag — skips cycle if previous run still active'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 11. ENVIRONMENT CONFIGURATION
  // ═══════════════════════════════════════════════════════════════
  h1('11. Environment Configuration'),
  h2('11.1 Backend .env'),
  mkTable(['Variable','Value / Notes','Required'],[2800,5000,1560],[
    ['APP_DB_HOST / USER / PASSWORD / NAME','MySQL credentials for voltstartev_db','Required'],
    ['STEVE_DB_HOST / USER / PASSWORD','MySQL credentials for stevedb','Required'],
    ['JWT_SECRET','STh0J/t5Wwk2pNNTg4N11bsfNThieH3gkZft9m8gXAE= (change in prod)','Required'],
    ['STEVE_URL','http://localhost:8080','Required'],
    ['STEVE_API_USER / PASS','admin / admin (change in prod)','Required'],
    ['RAZORPAY_KEY_ID / KEY_SECRET','rzp_test_* keys for test mode','Required'],
    ['RAZORPAY_WEBHOOK_SECRET','Real secret from Razorpay Dashboard → Settings → Webhooks','Required'],
    ['ANTHROPIC_API_KEY','sk-ant-* key for Volt AI assistant','Required'],
    ['FLEET_ENABLED','true — enables fleet management APIs and UI','Optional'],
    ['SKIP_OTP','true — MVP mode, bypasses OTP flow (dev only)','Dev only'],
    ['PORT','3000','Optional (default 3000)'],
  ]),
  sp(8),

  h2('11.2 Known Pending Items'),
  mkTable(['Item','Severity','Action Required'],[2000,1200,6160],[
    ['HTTPS/SSL not configured','High','Add SSL certificate (Let\'s Encrypt) to GCP instance before production'],
    ['RAZORPAY_WEBHOOK_SECRET placeholder','Critical','Set real webhook secret from Razorpay Dashboard'],
    ['Rate limiting on auth endpoints','High','Add express-rate-limit to /login, /register, /otp/* routes'],
    ['charger_capabilities table not created','Critical','Run CREATE TABLE migration (see ops guide)'],
    ['charger_config missing location columns','High','Run ALTER TABLE migration (see ops guide)'],
    ['AsyncStorage for token (not encrypted)','Medium','Migrate to expo-secure-store in production build'],
    ['newArchEnabled: false in app.json','Low','Remove or set to true — required for production builds on Expo SDK 55+'],
    ['expo-notifications deprecation warnings','Low','Replace shouldShowAlert with shouldShowBanner + shouldShowList'],
    ['A1A2A3A4 missing from stevedb.user_ocpp_tag','High','INSERT user_ocpp_tag row (user_pk=33, ocpp_tag_pk=465)'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 12. DEPLOYMENT
  // ═══════════════════════════════════════════════════════════════
  h1('12. Deployment'),
  h2('12.1 Server Setup'),
  flowBox([
    'GCP Compute Engine — asia-south1 (Mumbai)',
    'IP: 136.113.7.146',
    '',
    'Ports:',
    '  3000  — VoltStartEV Backend (Express + WebSocket)',
    '  8080  — SteVe OCPP Server',
    '  3306  — MySQL 8.0 (internal only)',
    '',
    'Process Manager: PM2',
    '  pm2 start npm --name voltstartev -- run start',
    '  pm2 restart voltstartev',
    '  pm2 logs voltstartev --lines 100',
    '',
    'Build & Deploy:',
    '  cd /build/VoltStartEV_Backend',
    '  npx tsc                    # compile TypeScript',
    '  pm2 restart voltstartev    # zero-downtime restart',
    '',
    'SteVe:',
    '  Running at /build/steve',
    '  Java 21 process managed separately',
    '  Config: ws://0.0.0.0:8080/steve/websocket/CentralSystemService/{chargeBoxId}',
  ]),
  sp(12),

  h2('12.2 Mobile Build'),
  mkTable(['Build Type','Command','Use Case'],[2400,3800,3160],[
    ['Expo Go (dev)','npx expo start --clear','Development — no Android SDK needed, instant QR scan'],
    ['Development Build','eas build --profile development','Full native features: push notifications, secure store'],
    ['Production APK','eas build --profile production --platform android','Play Store distribution'],
    ['Production IPA','eas build --profile production --platform ios','App Store distribution'],
  ]),
  sp(8),
  note('Expo Go SDK 53 is the current dev workflow. PROVIDER_GOOGLE removed from MapView for Expo Go compatibility. For production builds, add PROVIDER_GOOGLE back and configure Google Maps API key in app.json.','warning'),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 13. KEY ARCHITECTURAL DECISIONS
  // ═══════════════════════════════════════════════════════════════
  h1('13. Key Architectural Decisions'),
  mkTable(['Decision','Choice','Rationale'],[2800,2400,4160],[
    ['WebSocket library','Native ws (no Socket.io)','Smaller bundle, direct control, Expo Go compatible'],
    ['WebSocket auth method','?token= URL query param','React Native WebSocket does not support custom headers on upgrade'],
    ['Payment UI','Razorpay WebView (not native SDK)','Razorpay native SDK incompatible with Expo Go SDK 53+'],
    ['Pricing stored at session creation','rate_per_kwh + session_fee in sessions table','Changing prices never breaks in-progress or historical billing'],
    ['Dual DB schema','stevedb (SteVe) + voltstartev_db (app)','SteVe owns its schema — app DB extends it without modifying SteVe'],
    ['No ORM','Raw mysql2 queries','Full control over query optimization, no ORM abstraction overhead'],
    ['parseInt(String(value), 10) everywhere','Required for all LIMIT/OFFSET','mysql2 driver rejects non-integer LIMIT params — recurring pattern'],
    ['Monotonic cost guard','safeCost = Math.max(prevCost, newCost)','Prevents billing decreases from out-of-order or backward meter values'],
    ['Atomic stop_requested flag','UPDATE ... WHERE stop_requested=0','Prevents duplicate RemoteStop when multiple MeterValues arrive simultaneously'],
    ['RFID soft delete','is_active=0 not DELETE','Preserves audit trail, SteVe sync handles actual blocking'],
    ['Map markers: pure View','BoltIcon (View) not SVG icon','SVG inside Marker causes invisible/misaligned markers on Android Fabric'],
    ['Session store in Zustand','ws, wsConnected, reconnectAttempts all in interface','Zustand ignores set() calls for undeclared fields — all WS state must be in interface'],
    ['filteredChargers useMemo','Per-connector maxPowerWatts for power filter','Charger-level maxPower is inaccurate for multi-connector chargers'],
  ]),
  sp(16),
  pb(),

  // ═══════════════════════════════════════════════════════════════
  // 14. FUTURE ROADMAP
  // ═══════════════════════════════════════════════════════════════
  h1('14. Future Roadmap'),
  h2('14.1 Short Term (Pre-Production)'),
  bul('SSL/TLS configuration on GCP instance'),
  bul('Rate limiting on all auth endpoints'),
  bul('Run charger_config and charger_capabilities DB migrations'),
  bul('Set real Razorpay webhook secret'),
  bul('Migrate token storage to expo-secure-store'),
  bul('OTP flow full implementation in mobile (SKIP_OTP=false)'),
  bul('SendLocalList — push tag list to chargers for offline RFID authorization'),
  sp(8),
  h2('14.2 Medium Term'),
  bul('TriggerMessage OCPP endpoint — force StatusNotification/MeterValues on demand'),
  bul('Smart charging load balancing — rebalanceConnectorPower() when multiple EVs connect to same charger'),
  bul('GetConfiguration / ChangeConfiguration — operator UI to configure charger parameters'),
  bul('Diagnostics integration — GetDiagnostics for fault investigation'),
  bul('Admin dashboard — web-based operator panel for charger management'),
  bul('Analytics dashboard — energy delivered, revenue, utilization rates'),
  sp(8),
  h2('14.3 Long Term'),
  bul('OCPP 2.0.1 migration — native smart charging, device management, improved security model'),
  bul('Real-time load balancing engine — dynamic power allocation across charger network'),
  bul('Vehicle-to-Grid (V2G) support for compatible chargers'),
  bul('Integration with energy providers for time-of-use pricing based on live grid rates'),
  bul('Multi-tenant operator portal — white-label platform for other EV charging networks'),
  sp(16),

];

// ── Build document ────────────────────────────────────────────────
const doc = new Document({
  numbering:{config:[
    {reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'•',
      alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]},
  ]},
  styles:{
    default:{document:{run:{font:'Arial',size:22}}},
    paragraphStyles:[
      {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:36,bold:true,font:'Arial'},paragraph:{spacing:{before:480,after:160},outlineLevel:0}},
      {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:28,bold:true,font:'Arial'},paragraph:{spacing:{before:320,after:120},outlineLevel:1}},
      {id:'Heading3',name:'Heading 3',basedOn:'Normal',next:'Normal',quickFormat:true,
        run:{size:24,bold:true,font:'Arial'},paragraph:{spacing:{before:200,after:80},outlineLevel:2}},
    ]},
  sections:[{
    properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1440,bottom:1440,left:1440}}},
    headers:{default:hdr()},
    footers:{default:ftr()},
    children,
  }],
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync('C:/voltstartEV/VoltStartEV_SystemDesign.docx', buf);
  console.log('Done');
}).catch(e=>{console.error(e);process.exit(1);});

