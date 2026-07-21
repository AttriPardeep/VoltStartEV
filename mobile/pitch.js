const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

// ── Icon helper ──────────────────────────────────────
const {
  FaBolt, FaMapMarkerAlt, FaCreditCard, FaChartLine, FaUsers,
  FaCar, FaMobileAlt, FaMicrophone, FaBuilding, FaHandshake,
  FaCheckCircle, FaRocket, FaShieldAlt, FaWifi, FaLeaf,
  FaIndustry, FaCoins, FaGlobe, FaStar, FaCog,
} = require("react-icons/fa");
const { MdElectricCar, MdSpeed, MdPayment, MdAnalytics,
  MdLocationOn, MdSupportAgent, MdAutoMode } = require("react-icons/md");

async function icon(Comp, color="#FFFFFF", size=256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ── Palette ──────────────────────────────────────────
const BG    = "060F1E"; // near-black
const TEAL  = "0D9488";
const TEAL2 = "14B8A6";
const TEAL3 = "5EEAD4";
const WHITE = "FFFFFF";
const GRAY  = "94A3B8";
const SLATE = "1E293B";
const CARD  = "0D1F35";
const AMBER = "F59E0B";
const GREEN = "10B981";
const PURP  = "7C3AED";

const mkShadow = () => ({ type:"outer", blur:8, offset:3, angle:135, color:"000000", opacity:0.25 });

// ── Slide factory helpers ────────────────────────────
function darkSlide(pres, bg=BG) {
  const s = pres.addSlide();
  s.background = { color: bg };
  return s;
}

function card(s, x,y,w,h, color=CARD) {
  s.addShape(pres_global.shapes.RECTANGLE, {
    x,y,w,h, fill:{color}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow(),
  });
}

function accent(s,x,y,h,color=TEAL) {
  s.addShape(pres_global.shapes.RECTANGLE, { x,y,w:0.06,h, fill:{color}, line:{type:"none"} });
}

function tag(s, text, x, y, bg=TEAL, tc=BG) {
  const w = text.length * 0.095 + 0.25;
  s.addShape(pres_global.shapes.ROUNDED_RECTANGLE, {
    x,y,w,h:0.28, fill:{color:bg}, line:{type:"none"}, rectRadius:0.05,
  });
  s.addText(text, { x,y,w,h:0.28, fontSize:9, color:tc, bold:true, align:"center", valign:"middle", margin:0 });
}

let pres_global;

async function build() {
  const pres = new pptxgen();
  pres_global = pres;
  pres.layout  = "LAYOUT_16x9";
  pres.author  = "VoltStartEV";
  pres.title   = "VoltStartEV — Investor & Client Pitch";

  // ── Pre-render icons ─────────────────────────────────
  const IC = {
    bolt:    await icon(FaBolt,    TEAL2),
    map:     await icon(FaMapMarkerAlt, TEAL2),
    card:    await icon(FaCreditCard,   TEAL2),
    chart:   await icon(FaChartLine,    TEAL2),
    users:   await icon(FaUsers,        TEAL2),
    car:     await icon(FaCar,          TEAL2),
    mobile:  await icon(FaMobileAlt,    TEAL2),
    mic:     await icon(FaMicrophone,   AMBER),
    bldg:    await icon(FaBuilding,     TEAL2),
    shake:   await icon(FaHandshake,    TEAL2),
    check:   await icon(FaCheckCircle,  GREEN),
    rocket:  await icon(FaRocket,       AMBER),
    shield:  await icon(FaShieldAlt,    TEAL2),
    wifi:    await icon(FaWifi,         TEAL2),
    leaf:    await icon(FaLeaf,         GREEN),
    coins:   await icon(FaCoins,        AMBER),
    globe:   await icon(FaGlobe,        TEAL2),
    star:    await icon(FaStar,         AMBER),
    cog:     await icon(FaCog,          TEAL2),
    speed:   await icon(MdSpeed,        TEAL2),
    pay:     await icon(MdPayment,      TEAL2),
    analyt:  await icon(MdAnalytics,    TEAL2),
    support: await icon(MdSupportAgent, TEAL2),
    auto:    await icon(MdAutoMode,     AMBER),
    eccar:   await icon(MdElectricCar,  TEAL2),
    // white variants for dark cards
    boltW:   await icon(FaBolt,    WHITE),
    mapW:    await icon(FaMapMarkerAlt, WHITE),
    cardW:   await icon(FaCreditCard,   WHITE),
    chartW:  await icon(FaChartLine,    WHITE),
    usersW:  await icon(FaUsers,        WHITE),
    shieldW: await icon(FaShieldAlt,    WHITE),
    rocketW: await icon(FaRocket,       WHITE),
    leafW:   await icon(FaLeaf,         WHITE),
    coinsW:  await icon(FaCoins,        WHITE),
    checkW:  await icon(FaCheckCircle,  WHITE),
  };

  // ══════════════════════════════════════════════════════
  // SLIDE 1 — COVER
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    // Deep gradient effect with layered shapes
    s.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:10,h:5.625, fill:{color:"0A1628"}, line:{type:"none"} });
    s.addShape(pres.shapes.OVAL, { x:-1,y:-1,w:5,h:5, fill:{color:TEAL,transparency:92}, line:{type:"none"} });
    s.addShape(pres.shapes.OVAL, { x:7,y:2,w:4,h:4, fill:{color:PURP,transparency:92}, line:{type:"none"} });

    // Bolt icon large
    s.addImage({ data:IC.bolt, x:0.55, y:1.1, w:0.75, h:0.75 });

    // VoltStartEV wordmark
    s.addText("VoltStart", { x:1.35,y:1.0,w:5,h:0.9, fontSize:44, color:WHITE, bold:true, fontFace:"Trebuchet MS", margin:0 });
    s.addText("EV", { x:5.6, y:1.0, w:1.2, h:0.9, fontSize:44, color:TEAL2, bold:true, fontFace:"Trebuchet MS", margin:0 });

    // Tagline
    s.addText("Smart. Scalable. Sustainable Charging.", {
      x:0.55,y:1.95,w:8,h:0.5, fontSize:17, color:TEAL3, italic:true, fontFace:"Calibri",
    });

    // Divider line
    s.addShape(pres.shapes.LINE, { x:0.55,y:2.6,w:8.9,h:0, line:{color:TEAL,width:0.5} });

    // Sub headline
    s.addText("Platform Overview  ·  Business Case  ·  Customization  ·  Future Vision", {
      x:0.55,y:2.75,w:9,h:0.4, fontSize:13, color:GRAY, align:"center",
    });

    // Stat pills
    const stats = [["12+","Charger Types"],["5","Pricing Models"],["AI","Powered"],["OCPP","1.6J"]];
    stats.forEach(([val,lbl],i) => {
      const x = 0.55 + i*2.35;
      s.addShape(pres.shapes.RECTANGLE, { x,y:3.3,w:2.1,h:0.9, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addText(val, { x,y:3.3,w:2.1,h:0.5, fontSize:22, color:TEAL2, bold:true, align:"center", valign:"bottom" });
      s.addText(lbl, { x,y:3.7,w:2.1,h:0.4, fontSize:10, color:GRAY, align:"center", valign:"top" });
    });

    s.addText("Confidential — For Discussion Only", {
      x:0,y:5.3,w:10,h:0.3, fontSize:9, color:"334155", align:"center",
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 2 — THE PROBLEM WE SOLVE
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:10,h:5.625, fill:{color:"060F1E"}, line:{type:"none"} });

    s.addText("THE PROBLEM", { x:0.5,y:0.3,w:4,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("EV Charging is Broken\nfor Drivers & Operators", {
      x:0.5,y:0.6,w:5.5,h:1.3, fontSize:30, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    const probs = [
      ["No unified app","Drivers juggle multiple apps per network — no single experience"],
      ["No real-time data","Maps show chargers that are offline, in-use, or broken"],
      ["Opaque billing","Users don't know cost until after charging — no live estimates"],
      ["No smart control","Operators can't manage power, pricing, or reservations remotely"],
      ["No customisation","Builders & societies get a one-size-fits-all solution that fits nobody"],
    ];
    probs.forEach(([ttl,desc],i) => {
      const y = 1.9 + i*0.65;
      s.addShape(pres.shapes.RECTANGLE, { x:0.5,y:y-0.03,w:5.8,h:0.58, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5} });
      s.addShape(pres.shapes.RECTANGLE, { x:0.5,y:y-0.03,w:0.06,h:0.58, fill:{color:"EF4444"}, line:{type:"none"} });
      s.addText(ttl, { x:0.7,y:y,w:2,h:0.25, fontSize:11, color:WHITE, bold:true, margin:0 });
      s.addText(desc, { x:0.7,y:y+0.24,w:5.4,h:0.25, fontSize:10, color:GRAY, margin:0 });
    });

    // Right side — market pain
    s.addShape(pres.shapes.RECTANGLE, { x:6.6,y:1.3,w:3.2,h:4.0, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
    s.addImage({ data:IC.eccar, x:7.3,y:1.5,w:1.8,h:1.1 });
    s.addText("India's EV Market", { x:6.7,y:2.65,w:3,h:0.4, fontSize:13, color:TEAL2, bold:true, align:"center" });

    const mstats = [["10M+","EVs on road by 2030"],["₹3.7L Cr","Market size projection"],["<5%","Chargers with smart mgmt"]];
    mstats.forEach(([v,l],i) => {
      const y = 3.15 + i*0.7;
      s.addText(v, { x:6.7,y,w:3,h:0.35, fontSize:20, color:AMBER, bold:true, align:"center" });
      s.addText(l, { x:6.7,y:y+0.32,w:3,h:0.28, fontSize:10, color:GRAY, align:"center" });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 3 — WHAT IS VOLTSTART EV
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addText("OUR SOLUTION", { x:0.5,y:0.3,w:5,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("One Platform. Every Charger. Every Driver.", {
      x:0.5,y:0.6,w:9,h:0.8, fontSize:28, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    // Three pillars
    const pillars = [
      { ic:IC.mobile,  t:"Mobile App",    d:"React Native iOS & Android app for drivers — discover, charge, pay, track in one place." },
      { ic:IC.cog,     t:"Smart Backend", d:"Node.js + TypeScript API on GCP — real-time OCPP 1.6J, WebSocket telemetry, AI-powered." },
      { ic:IC.bldg,    t:"Operator Portal",d:"Charger management, dynamic pricing, fleet control, reporting — all from a dashboard." },
    ];
    pillars.forEach(({ic,t,d},i) => {
      const x = 0.4 + i * 3.2;
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.5,w:3.0,h:3.8, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      // accent top
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.5,w:3.0,h:0.07, fill:{color:TEAL}, line:{type:"none"} });
      s.addImage({ data:ic, x:x+1.2,y:1.75, w:0.6,h:0.6 });
      s.addText(t, { x,y:2.45,w:3,h:0.4, fontSize:14, color:WHITE, bold:true, align:"center" });
      s.addText(d, { x:x+0.15,y:2.9,w:2.7,h:1.5, fontSize:11, color:GRAY, align:"center", wrap:true });
    });

    // Tech badges
    ["React Native","Node.js + TypeScript","MySQL 8","OCPP 1.6J","GCP Mumbai","Razorpay","Claude AI"].forEach((t,i)=>{
      const x = 0.4 + (i%4)*2.4;
      const y = i<4 ? 5.05 : 5.05;
      tag(s,t,x,5.15, "0D1F35", TEAL3);
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 4 — DRIVER EXPERIENCE
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addText("DRIVER EXPERIENCE", { x:0.5,y:0.25,w:5,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("Everything a Driver Needs\nin One App", {
      x:0.5,y:0.55,w:5.5,h:1.0, fontSize:26, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    const feats = [
      [IC.map,    "Live Charger Map",     "Real-time availability, filter by power, connector type, price, distance"],
      [IC.bolt,   "One-Tap Charging",     "Start & stop sessions via app or RFID card — no juggling apps"],
      [IC.coins,  "Live Cost Tracking",   "See cost and energy updating every 5 seconds during charging"],
      [IC.card,   "Prepaid Wallet",       "Razorpay top-up — session auto-stops if balance runs low"],
      [IC.car,    "Vehicle Profiles",     "Store EV details, set target SOC — auto-stop at your desired charge %"],
      [IC.chart,  "Session History",      "Full charging history with cost, energy, duration, pricing breakdown"],
      [IC.shield, "RFID Cards",           "Register physical RFID card — tap-to-charge without opening app"],
      [IC.bldg,   "Fleet Management",     "Fleet admin controls, billing mode, monthly limits per driver"],
    ];
    feats.forEach(([ic,t,d],i) => {
      const col = i%2, row = Math.floor(i/2);
      const x = 0.4 + col*4.9;
      const y = 1.65 + row*0.92;
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:4.7,h:0.8, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5} });
      s.addImage({ data:ic, x:x+0.15,y:y+0.17, w:0.42,h:0.42 });
      s.addText(t, { x:x+0.68,y:y+0.08,w:3.8,h:0.3, fontSize:12, color:WHITE, bold:true, margin:0 });
      s.addText(d, { x:x+0.68,y:y+0.37,w:3.8,h:0.35, fontSize:9.5, color:GRAY, margin:0 });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 5 — AI FEATURES
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    // Gradient glow
    s.addShape(pres.shapes.OVAL, { x:3,y:-1,w:6,h:5, fill:{color:PURP,transparency:88}, line:{type:"none"} });

    s.addText("AI-POWERED", { x:0.5,y:0.25,w:5,h:0.35, fontSize:10, color:AMBER, bold:true, charSpacing:4 });
    s.addText("Intelligence Built Into\nEvery Charging Interaction", {
      x:0.5,y:0.55,w:6,h:1.1, fontSize:27, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    // Volt AI card — left large
    s.addShape(pres.shapes.RECTANGLE, { x:0.4,y:1.75,w:4.5,h:3.5, fill:{color:"0D1526"}, line:{color:PURP,width:1}, shadow:mkShadow() });
    s.addImage({ data:IC.bolt, x:1.6,y:1.95,w:0.6,h:0.6 });
    s.addText("Volt AI", { x:2.15,y:1.93,w:2,h:0.5, fontSize:22, color:WHITE, bold:true, margin:0 });
    s.addText("Powered by Anthropic Claude", { x:0.5,y:2.5,w:4.3,h:0.35, fontSize:10, color:GRAY, align:"center" });
    const voltFeats = [
      "Ask Volt anything about your charging session",
      "Get AI-generated monthly charging reports",
      "Personalized cost optimization tips",
      "Fleet usage insights and anomaly alerts",
      "Natural language charger search",
    ];
    voltFeats.forEach((f,i) => {
      s.addImage({ data:IC.checkW, x:0.6,y:2.98+i*0.46, w:0.22,h:0.22 });
      s.addText(f, { x:0.9,y:2.95+i*0.46,w:3.8,h:0.35, fontSize:10.5, color:"CBD5E1", margin:0 });
    });

    // Right — upcoming AI features
    const aiUpcoming = [
      { ic:IC.mic, t:"Voice Wake-Up", d:"Next Release — 'Hey Volt, start charging my car' — hands-free session control", badge:"COMING SOON", bc:AMBER },
      { ic:IC.auto, t:"Smart Auto-Stop", d:"AI predicts optimal charge endpoint based on vehicle, trip schedule, and price", badge:"IN APP NOW", bc:GREEN },
      { ic:IC.analyt, t:"Predictive Pricing", d:"AI recommends best time to charge based on live ToU rates and your usage patterns", badge:"ROADMAP", bc:PURP },
    ];
    aiUpcoming.forEach(({ic,t,d,badge,bc},i) => {
      const y = 1.75 + i*1.2;
      s.addShape(pres.shapes.RECTANGLE, { x:5.2,y,w:4.5,h:1.05, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addImage({ data:ic, x:5.35,y:y+0.27, w:0.45,h:0.45 });
      s.addText(t, { x:5.9,y:y+0.1,w:2.6,h:0.35, fontSize:13, color:WHITE, bold:true, margin:0 });
      tag(s,badge, 8.1, y+0.1, bc, bc===AMBER?BG:WHITE);
      s.addText(d, { x:5.9,y:y+0.48,w:3.7,h:0.45, fontSize:9.5, color:GRAY, margin:0, wrap:true });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 6 — CUSTOMISATION
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addText("FULLY CUSTOMISABLE", { x:0.5,y:0.25,w:6,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("Your Brand. Your Rules.\nYour Charging Network.", {
      x:0.5,y:0.55,w:7,h:1.1, fontSize:27, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    // Three audience columns
    const audiences = [
      {
        title:"Builders & Societies",
        ic: IC.bldg,
        color: TEAL,
        pts:[
          "Your building logo on the app",
          "Resident-only access (private chargers)",
          "Branded charging screen & receipts",
          "Flat monthly resident pricing",
          "Usage reports per unit/flat",
          "Admin dashboard for facility manager",
        ],
      },
      {
        title:"Corporate Campuses",
        ic: IC.users,
        color: AMBER,
        pts:[
          "Employee RFID card integration",
          "Cost centre billing per department",
          "Fleet vehicle management",
          "Peak/off-peak pricing by shift",
          "GST-compliant monthly invoices",
          "EV policy compliance reporting",
        ],
      },
      {
        title:"Charging Operators",
        ic: IC.coins,
        color: PURP,
        pts:[
          "Multi-site network management",
          "Dynamic pricing per charger",
          "Public + private charger mix",
          "Revenue sharing configuration",
          "White-label app option",
          "API access for third-party integration",
        ],
      },
    ];
    audiences.forEach(({title,ic,color,pts},i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.7,w:3.1,h:3.75, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.7,w:3.1,h:0.07, fill:{color}, line:{type:"none"} });
      s.addImage({ data:ic, x:x+1.3,y:1.85, w:0.5,h:0.5 });
      s.addText(title, { x,y:2.42,w:3.1,h:0.4, fontSize:12, color:WHITE, bold:true, align:"center" });
      s.addText(pts.map(p=>`• ${p}`).join("\n"), {
        x:x+0.15,y:2.88,w:2.8,h:2.4, fontSize:9.5, color:GRAY, wrap:true,
      });
    });

    // Bottom note
    s.addShape(pres.shapes.RECTANGLE, { x:0.3,y:5.28,w:9.4,h:0.3, fill:{color:"0D1526"}, line:{color:"1E3A5F",width:0.5} });
    s.addText("⚡  Builder & society logos embedded in app  ·  Custom branding on receipts, notifications, and login screen  ·  Configurable per deployment", {
      x:0.4,y:5.3,w:9.2,h:0.26, fontSize:9, color:TEAL3, align:"center",
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 7 — PRICING & REVENUE MODELS
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addText("REVENUE MODELS", { x:0.5,y:0.25,w:5,h:0.35, fontSize:10, color:AMBER, bold:true, charSpacing:4 });
    s.addText("Flexible Monetisation for Every Agreement", {
      x:0.5,y:0.58,w:9,h:0.7, fontSize:26, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    // Pricing model cards — top row
    const models = [
      { t:"Per Unit (kWh)",     d:"Charge exactly for energy delivered. Standard for most public chargers.", eg:"₹9 / kWh", color:TEAL },
      { t:"Per Transaction",    d:"Fixed fee per charging session regardless of energy or time.", eg:"₹50 / session", color:AMBER },
      { t:"Per Minute",         d:"Time-based billing. Best for destination chargers & parking.", eg:"₹2 / minute", color:GREEN },
      { t:"Tiered Power",       d:"Rate changes with charging speed. HPC gets premium rate.", eg:"₹16–₹28 / kWh", color:PURP },
      { t:"Time of Use",        d:"Peak/off-peak/shoulder rates by hour. Shifts load to cheap periods.", eg:"₹6–₹14 / kWh", color:"3B82F6" },
      { t:"Free / Complimentary",d:"No charge to user. For hotels, offices, incentive programmes.", eg:"₹0", color:GRAY },
    ];
    models.forEach(({t,d,eg,color},i) => {
      const col=i%3, row=Math.floor(i/3);
      const x = 0.35 + col*3.2, y = 1.45 + row*1.55;
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:3.05,h:1.4, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:0.06,h:1.4, fill:{color}, line:{type:"none"} });
      s.addText(t,  { x:x+0.18,y:y+0.1, w:2.6,h:0.35, fontSize:12, color:WHITE, bold:true, margin:0 });
      s.addText(d,  { x:x+0.18,y:y+0.47,w:2.6,h:0.55, fontSize:9.5, color:GRAY, margin:0, wrap:true });
      s.addText(eg, { x:x+0.18,y:y+1.05,w:2.6,h:0.28, fontSize:11, color, bold:true, margin:0 });
    });

    // Right side — VoltStartEV revenue
    s.addShape(pres.shapes.RECTANGLE, { x:9.7,y:1.45,w:0.01,h:3,fill:{color:CARD},line:{type:"none"} }); // spacer
    // Revenue share note
    s.addShape(pres.shapes.RECTANGLE, { x:0.35,y:4.55,w:9.35,h:0.9, fill:{color:"0D1526"}, line:{color:TEAL,width:0.5} });
    s.addText("VoltStartEV Revenue Share", { x:0.5,y:4.6,w:3,h:0.35, fontSize:11, color:TEAL2, bold:true });
    s.addText("Platform fee: 5–10% of transaction value  ·  SaaS subscription for operators  ·  Custom integration fee for white-label deployments  ·  Revenue sharing negotiated per agreement", {
      x:0.5,y:4.9,w:9.2,h:0.45, fontSize:10, color:GRAY, wrap:true,
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 8 — REAL-TIME TECH
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addText("TECHNOLOGY", { x:0.5,y:0.25,w:5,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("Real-Time Architecture\nBuilt for Scale", {
      x:0.5,y:0.55,w:6,h:1.0, fontSize:27, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    // Flow diagram (text-based)
    s.addShape(pres.shapes.RECTANGLE, { x:0.4,y:1.65,w:9.2,h:3.7, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5} });

    // Row 1: Mobile → Backend → SteVe → Charger
    const nodes = [
      {t:"Mobile App",   s:"React Native",    x:0.6,  color:TEAL},
      {t:"Backend API",  s:"Node.js + TS",    x:3.0,  color:TEAL},
      {t:"SteVe OCPP",   s:"Java / MySQL",    x:5.5,  color:AMBER},
      {t:"EV Charger",   s:"OCPP 1.6J",       x:8.0,  color:GREEN},
    ];
    nodes.forEach(({t,s:sub,x,color}) => {
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.9,w:1.9,h:0.9, fill:{color:"0A1628"}, line:{color,width:0.8} });
      s.addText(t,   { x,y:1.92,w:1.9,h:0.4, fontSize:11, color:WHITE, bold:true, align:"center" });
      s.addText(sub, { x,y:2.3, w:1.9,h:0.4, fontSize:9,  color:GRAY,             align:"center" });
    });
    // Arrows
    ["REST API  +  WebSocket","OCPP Webhooks","OCPP 1.6J WS"].forEach((lbl,i) => {
      const ax = 2.5 + i*2.5;
      s.addShape(pres.shapes.LINE, { x:ax,y:2.35,w:0.5,h:0, line:{color:TEAL,width:1.5} });
      s.addText("→", { x:ax,y:2.25,w:0.5,h:0.25, fontSize:14, color:TEAL, align:"center" });
      s.addText(lbl, { x:ax-0.2,y:2.58,w:0.9,h:0.3, fontSize:7.5, color:GRAY, align:"center" });
    });

    // Key capabilities row
    const caps = [
      [IC.wifi,   "WebSocket",    "Real-time telemetry every 5 seconds — no polling"],
      [IC.speed,  "< 200ms",      "API response time with in-memory cache layer"],
      [IC.shield, "Atomic Ops",   "Race-condition-safe billing and auto-stop logic"],
      [IC.analyt, "Live Cost",    "Cost updates every meter value during session"],
    ];
    caps.forEach(({0:ic,1:t,2:d},i) => {
      const x = 0.55 + i*2.3;
      s.addShape(pres.shapes.RECTANGLE, { x,y:3.05,w:2.15,h:2.05, fill:{color:"0A1628"}, line:{color:"1E3A5F",width:0.5} });
      s.addImage({ data:ic, x:x+0.83,y:3.2, w:0.5,h:0.5 });
      s.addText(t, { x,y:3.78,w:2.15,h:0.35, fontSize:12, color:TEAL2, bold:true, align:"center" });
      s.addText(d, { x:x+0.1,y:4.15,w:1.95,h:0.8, fontSize:9, color:GRAY, align:"center", wrap:true });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 9 — FUTURE ROADMAP
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addShape(pres.shapes.OVAL, { x:5,y:1,w:7,h:6, fill:{color:TEAL,transparency:94}, line:{type:"none"} });

    s.addText("ROADMAP", { x:0.5,y:0.25,w:4,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("What's Coming", { x:0.5,y:0.55,w:5,h:0.7, fontSize:30, color:WHITE, bold:true, fontFace:"Trebuchet MS" });

    const phases = [
      {
        phase:"Now — Live",
        color:GREEN,
        items:[
          "OCPP 1.6J compliant charging",
          "5 pricing models (per kWh, tiered, ToU, per-min, free)",
          "Real-time WebSocket telemetry",
          "Wallet + Razorpay payment",
          "RFID card management",
          "Fleet management + multi-vehicle",
          "Volt AI assistant (Claude)",
          "Push notifications",
          "SOC auto-stop + low-balance auto-stop",
        ],
      },
      {
        phase:"Next — 3 Months",
        color:AMBER,
        items:[
          "Voice wake-up: 'Hey Volt, start charging'",
          "Builder / society logo branding",
          "OTP-based authentication (production)",
          "HTTPS / SSL security hardening",
          "Smart charging load balancing",
          "SendLocalList (offline RFID auth)",
          "Admin web dashboard (operators)",
        ],
      },
      {
        phase:"Future — 6–12 Months",
        color:PURP,
        items:[
          "OCPP 2.0.1 migration",
          "V2G (Vehicle-to-Grid) support",
          "Dynamic grid-aware pricing",
          "Multi-tenant white-label platform",
          "Real-time load balancing engine",
          "Integration with energy providers",
          "Predictive AI charging recommendations",
        ],
      },
    ];
    phases.forEach(({phase,color,items},i) => {
      const x = 0.3 + i*3.25;
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.35,w:3.1,h:4.05, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.35,w:3.1,h:0.06, fill:{color}, line:{type:"none"} });
      s.addShape(pres.shapes.RECTANGLE, { x,y:1.35,w:3.1,h:0.5, fill:{color,transparency:85}, line:{type:"none"} });
      s.addText(phase, { x,y:1.38,w:3.1,h:0.45, fontSize:12, color:WHITE, bold:true, align:"center" });
      s.addText(items.map(p=>`• ${p}`).join("\n"), {
        x:x+0.15,y:1.92,w:2.8,h:3.3, fontSize:9.5, color:GRAY, wrap:true,
      });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 10 — SUSTAINABILITY
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addShape(pres.shapes.OVAL, { x:-2,y:1,w:7,h:6, fill:{color:GREEN,transparency:93}, line:{type:"none"} });

    s.addText("SUSTAINABILITY", { x:0.5,y:0.25,w:5,h:0.35, fontSize:10, color:GREEN, bold:true, charSpacing:4 });
    s.addText("Every Charge. Every kWh.\nMeasured and Reported.", {
      x:0.5,y:0.58,w:9,h:1.0, fontSize:26, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    const scards = [
      { ic:IC.leaf,  t:"Carbon Tracking",      d:"Every session records kWh and estimated CO₂ offset vs petrol. Show your ESG impact in real numbers." },
      { ic:IC.chart, t:"Energy Reports",        d:"Monthly PDF reports — total energy delivered, peak demand, cost savings vs fossil fuel per user." },
      { ic:IC.globe, t:"Grid-Aware Charging",   d:"Time-of-use pricing incentivises off-peak charging — reduces grid load during peak demand hours." },
      {
        ic: IC.analyt,
        t: "Fleet ESG Dashboard",
        d: "Corporate ESG reporting — fleet electrification progress, avoided emissions, renewable energy %"
      }
    ];
    scards.forEach(({ic,t,d},i) => {
      const col=i%2, row=Math.floor(i/2);
      const x = 0.35 + col*4.85, y = 1.7 + row*1.85;
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:4.6,h:1.65, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:0.06,h:1.65, fill:{color:GREEN}, line:{type:"none"} });
      s.addImage({ data:ic, x:x+0.18,y:y+0.6, w:0.45,h:0.45 });
      s.addText(t, { x:x+0.75,y:y+0.15,w:3.6,h:0.4, fontSize:13, color:WHITE, bold:true, margin:0 });
      s.addText(d, { x:x+0.75,y:y+0.55,w:3.6,h:0.95, fontSize:10, color:GRAY, wrap:true, margin:0 });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 11 — WHY VOLTSTART EV
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres);
    s.addText("WHY VOLTSTART EV", { x:0.5,y:0.25,w:6,h:0.35, fontSize:10, color:TEAL2, bold:true, charSpacing:4 });
    s.addText("Built Different.\nDeployed Fast.", {
      x:0.5,y:0.55,w:6,h:1.0, fontSize:28, color:WHITE, bold:true, fontFace:"Trebuchet MS",
    });

    const diff = [
      { ic:IC.rocket, t:"Production-Ready",   d:"Not a prototype. Running live sessions across multiple charger types with real billing and real users.", color:AMBER },
      { ic:IC.shield, t:"OCPP 1.6J Compliant",d:"Standard protocol means compatibility with any OCPP-certified charger hardware from any vendor.", color:TEAL },
      { ic:IC.users,  t:"Multi-Tenant",       d:"One platform, unlimited deployments. Each client gets isolated data, branding, and pricing.", color:GREEN },
      { ic:IC.coins,  t:"Flexible Revenue",   d:"Six charging pricing models. Revenue sharing, SaaS, white-label — any agreement structure.", color:PURP },
      { ic:IC.mobile, t:"Native Mobile",       d:"React Native app — single codebase for iOS and Android. OTA updates without app store approval.", color:TEAL },
      { ic:IC.star,   t:"AI-First Design",    d:"Claude AI assistant baked into the app from day one — not bolted on later.", color:AMBER },
    ];
    diff.forEach(({ic,t,d,color},i) => {
      const col=i%2, row=Math.floor(i/3);
      const gridCol = i%3, gridRow = Math.floor(i/3);
      const x = 0.35 + gridCol*3.2, y = 1.65 + gridRow*1.85;
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:3.05,h:1.65, fill:{color:CARD}, line:{color:"1E3A5F",width:0.5}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x,y,w:3.05,h:0.06, fill:{color}, line:{type:"none"} });
      s.addImage({ data:ic, x:x+0.15,y:y+0.2, w:0.42,h:0.42 });
      s.addText(t, { x:x+0.68,y:y+0.15,w:2.2,h:0.38, fontSize:12, color:WHITE, bold:true, margin:0 });
      s.addText(d, { x:x+0.15,y:y+0.7, w:2.75,h:0.82, fontSize:9.5, color:GRAY, wrap:true, margin:0 });
    });
  }

  // ══════════════════════════════════════════════════════
  // SLIDE 12 — CLOSING / CTA
  // ══════════════════════════════════════════════════════
  {
    const s = darkSlide(pres, "060F1E");
    s.addShape(pres.shapes.OVAL, { x:-1,y:-1,w:6,h:6, fill:{color:TEAL,transparency:91}, line:{type:"none"} });
    s.addShape(pres.shapes.OVAL, { x:6,y:2,w:5,h:5, fill:{color:PURP,transparency:91}, line:{type:"none"} });

    // Bolt icon
    s.addImage({ data:IC.bolt, x:3.8,y:0.4, w:0.9,h:0.9 });

    s.addText("Let's Build India's", { x:1,y:1.35,w:8,h:0.6, fontSize:26, color:WHITE, fontFace:"Trebuchet MS", align:"center" });
    s.addText("Smartest Charging Network", { x:1,y:1.9,w:8,h:0.7, fontSize:30, color:TEAL2, bold:true, fontFace:"Trebuchet MS", align:"center" });
    s.addText("Together", { x:1,y:2.55,w:8,h:0.55, fontSize:22, color:GRAY, italic:true, align:"center" });

    // Contact / CTA cards
    const ctas = [
      { t:"Schedule a Demo",  d:"See a live charging session from the app", ic:IC.rocket },
      { t:"Custom Proposal",  d:"Tailored pricing & branding for your site", ic:IC.shake  },
      { t:"Pilot Programme",  d:"Deploy 2–3 chargers with full platform access", ic:IC.bolt  },
    ];
    ctas.forEach(({t,d,ic},i) => {
      const x = 0.85 + i*2.9;
      s.addShape(pres.shapes.RECTANGLE, { x,y:3.3,w:2.65,h:1.6, fill:{color:CARD}, line:{color:TEAL,width:0.8}, shadow:mkShadow() });
      s.addImage({ data:ic, x:x+1.1,y:3.42, w:0.45,h:0.45 });
      s.addText(t, { x,y:3.95,w:2.65,h:0.4, fontSize:12, color:WHITE, bold:true, align:"center" });
      s.addText(d, { x:x+0.1,y:4.35,w:2.45,h:0.45, fontSize:9.5, color:GRAY, align:"center", wrap:true });
    });

    s.addText("VoltStartEV  ·  pardeep.attri327@gmail.com  ·  github.com/AttriPardeep/VoltStartEV", {
      x:0,y:5.3,w:10,h:0.3, fontSize:9, color:"334155", align:"center",
    });
  }

  // ── Write file ────────────────────────────────────────
  await pres.writeFile({ fileName:"C:/voltstartEV/VoltStartEV_Pitch_Deck.pptx" });
  console.log("Done — VoltStartEV_Pitch_Deck.pptx");
}

build().catch(e=>{ console.error(e); process.exit(1); });






