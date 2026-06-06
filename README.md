<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Teachly – AI Lesson Note Generator</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/docx/8.5.0/docx.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.min.mjs" type="module"></script>
  <style>
    :root {
      --ink: #1a1207;
      --paper: #fdf8f0;
      --cream: #f5edd8;
      --amber: #e8a020;
      --amber-dark: #c4851a;
      --amber-glow: rgba(232,160,32,0.15);
      --rust: #c45c1a;
      --forest: #2d6a4f;
      --slate: #4a5568;
      --muted: #8a7e6e;
      --border: rgba(26,18,7,0.12);
      --shadow: 0 4px 24px rgba(26,18,7,0.10);
      --shadow-lg: 0 12px 48px rgba(26,18,7,0.16);
      --radius: 16px;
      --radius-sm: 10px;
    }
    * { margin:0; padding:0; box-sizing:border-box; }
    body {
      font-family: 'DM Sans', sans-serif;
      background: var(--paper);
      color: var(--ink);
      min-height: 100vh;
      overflow-x: hidden;
    }

    /* ── NAVBAR ── */
    nav {
      display: flex; align-items: center; justify-content: space-between;
      padding: 0 40px; height: 68px;
      background: rgba(253,248,240,0.92);
      backdrop-filter: blur(12px);
      border-bottom: 1px solid var(--border);
      position: sticky; top:0; z-index:100;
    }
    .logo {
      font-family: 'Playfair Display', serif;
      font-size: 1.7rem; font-weight:900;
      background: linear-gradient(135deg, var(--amber-dark), var(--rust));
      -webkit-background-clip: text; -webkit-text-fill-color: transparent;
      background-clip: text;
      letter-spacing: -0.5px;
    }
    .nav-links { display:flex; gap:8px; }
    .nav-btn {
      padding: 8px 20px; border-radius:8px; border: none;
      font-family: 'DM Sans', sans-serif; font-size: 0.9rem;
      cursor:pointer; transition: all 0.2s;
    }
    .nav-btn.ghost { background:transparent; color:var(--slate); }
    .nav-btn.ghost:hover { background:var(--cream); }
    .nav-btn.primary {
      background: linear-gradient(135deg, var(--amber), var(--amber-dark));
      color: white; font-weight:600;
      box-shadow: 0 2px 8px rgba(232,160,32,0.4);
    }
    .nav-btn.primary:hover { transform:translateY(-1px); box-shadow: 0 4px 16px rgba(232,160,32,0.5); }

    /* ── PAGES ── */
    .page { display:none; }
    .page.active { display:block; }

    /* ── HERO ── */
    .hero {
      min-height: calc(100vh - 68px);
      display:grid; grid-template-columns: 1fr 1fr;
      align-items: center; gap:60px;
      padding: 60px 80px;
      position:relative; overflow:hidden;
    }
    .hero::before {
      content:'';
      position:absolute; top:-200px; right:-200px;
      width:700px; height:700px;
      background: radial-gradient(circle, rgba(232,160,32,0.12) 0%, transparent 70%);
      pointer-events:none;
    }
    .hero-badge {
      display:inline-flex; align-items:center; gap:8px;
      background:var(--amber-glow); border: 1px solid rgba(232,160,32,0.3);
      padding:6px 14px; border-radius:100px;
      font-size:0.82rem; font-weight:500; color:var(--amber-dark);
      margin-bottom:20px;
    }
    .hero-badge::before { content:'✦'; font-size:0.7rem; }
    h1 {
      font-family:'Playfair Display', serif;
      font-size: clamp(2.4rem, 4vw, 3.6rem);
      line-height:1.1; font-weight:900;
      letter-spacing:-1px; margin-bottom:20px;
    }
    h1 span { color: var(--amber-dark); }
    .hero-sub {
      font-size:1.1rem; color:var(--slate); line-height:1.7;
      max-width:480px; margin-bottom:32px;
    }
    .hero-cta {
      display:flex; gap:12px; flex-wrap:wrap;
    }
    .btn-lg {
      padding: 14px 32px; border-radius:12px; border:none;
      font-family:'DM Sans',sans-serif; font-size:1rem; font-weight:600;
      cursor:pointer; transition:all 0.25s;
    }
    .btn-primary {
      background: linear-gradient(135deg, var(--amber), var(--amber-dark));
      color:white;
      box-shadow: 0 4px 20px rgba(232,160,32,0.45);
    }
    .btn-primary:hover { transform:translateY(-2px); box-shadow: 0 8px 28px rgba(232,160,32,0.55); }
    .btn-outline {
      background:transparent; color:var(--ink);
      border: 1.5px solid var(--border);
    }
    .btn-outline:hover { background:var(--cream); border-color:var(--amber); }

    .hero-visual {
      background: var(--cream);
      border-radius: 24px;
      border: 1px solid var(--border);
      padding:28px;
      box-shadow: var(--shadow-lg);
      position:relative;
    }
    .mock-header {
      display:flex; align-items:center; gap:8px; margin-bottom:20px;
    }
    .mock-dot { width:10px; height:10px; border-radius:50%; }
    .mock-week-list { display:flex; flex-direction:column; gap:8px; }
    .mock-week {
      display:flex; align-items:center; gap:12px;
      padding:10px 14px; background:white;
      border-radius:10px; border:1px solid var(--border);
      font-size:0.85rem;
    }
    .mock-week.active { border-color:var(--amber); background:var(--amber-glow); }
    .mock-week-num {
      width:24px; height:24px; border-radius:6px;
      background: var(--amber); color:white;
      font-size:0.75rem; font-weight:700;
      display:flex; align-items:center; justify-content:center;
    }
    .mock-week:not(.active) .mock-week-num { background:var(--cream); color:var(--muted); }
    .mock-status {
      margin-left:auto; font-size:0.72rem;
      padding:2px 8px; border-radius:100px;
    }
    .mock-status.done { background:#d1fae5; color:#065f46; }
    .mock-status.gen { background:var(--amber-glow); color:var(--amber-dark); }

    /* ── STATS ── */
    .stats { display:flex; gap:40px; margin-top:40px; padding-top:40px; border-top:1px solid var(--border); }
    .stat-num { font-family:'Playfair Display',serif; font-size:2rem; font-weight:900; color:var(--amber-dark); }
    .stat-label { font-size:0.82rem; color:var(--muted); margin-top:2px; }

    /* ── APP PAGE ── */
    .app-layout {
      display:grid; grid-template-columns: 300px 1fr;
      min-height: calc(100vh - 68px);
    }
    .sidebar {
      background: white;
      border-right: 1px solid var(--border);
      padding:24px;
      overflow-y:auto;
    }
    .sidebar-title {
      font-size:0.72rem; font-weight:600; letter-spacing:0.1em;
      text-transform:uppercase; color:var(--muted);
      margin-bottom:16px;
    }
    .main-content { padding:32px; overflow-y:auto; }

    /* ── UPLOAD ZONE ── */
    .upload-zone {
      border: 2px dashed var(--border);
      border-radius: var(--radius);
      padding: 48px 32px;
      text-align:center; cursor:pointer;
      transition: all 0.25s;
      background: white;
    }
    .upload-zone:hover, .upload-zone.drag-over {
      border-color: var(--amber);
      background: var(--amber-glow);
    }
    .upload-icon {
      width:56px; height:56px; margin: 0 auto 16px;
      background: var(--cream); border-radius:14px;
      display:flex; align-items:center; justify-content:center;
      font-size:1.6rem;
    }
    .upload-title { font-size:1rem; font-weight:600; margin-bottom:6px; }
    .upload-sub { font-size:0.85rem; color:var(--muted); }
    .upload-types { display:flex; gap:8px; justify-content:center; margin-top:16px; flex-wrap:wrap; }
    .type-badge {
      padding:3px 10px; border-radius:6px;
      background:var(--cream); font-size:0.78rem;
      font-weight:500; color:var(--slate);
    }

    /* ── FORM ── */
    .form-grid { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:24px; }
    .form-group { display:flex; flex-direction:column; gap:6px; }
    .form-group.full { grid-column:1/-1; }
    label { font-size:0.82rem; font-weight:600; color:var(--slate); }
    input, select, textarea {
      padding:10px 14px; border-radius:10px;
      border: 1.5px solid var(--border);
      font-family:'DM Sans',sans-serif; font-size:0.92rem;
      background:white; color:var(--ink);
      transition:border-color 0.2s;
      outline:none;
    }
    input:focus, select:focus, textarea:focus { border-color:var(--amber); }
    textarea { resize:vertical; min-height:120px; }

    /* ── PROGRESS ── */
    .progress-bar-wrap {
      background:var(--cream); border-radius:100px;
      height:8px; overflow:hidden; margin:8px 0;
    }
    .progress-bar-fill {
      height:100%;
      background: linear-gradient(90deg, var(--amber), var(--rust));
      border-radius:100px;
      transition:width 0.4s ease;
    }

    /* ── WEEK LIST ── */
    .week-item {
      display:flex; align-items:center; gap:10px;
      padding:10px 12px; border-radius:10px;
      cursor:pointer; transition:all 0.2s;
      margin-bottom:4px;
    }
    .week-item:hover { background:var(--cream); }
    .week-item.active { background:var(--amber-glow); }
    .week-badge {
      width:28px; height:28px; border-radius:8px;
      background:var(--cream); color:var(--muted);
      font-size:0.78rem; font-weight:700;
      display:flex; align-items:center; justify-content:center;
      flex-shrink:0;
    }
    .week-item.active .week-badge { background:var(--amber); color:white; }
    .week-info { flex:1; min-width:0; }
    .week-name { font-size:0.85rem; font-weight:500; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .week-status { font-size:0.72rem; color:var(--muted); }
    .week-status.done { color:var(--forest); }
    .week-status.gen { color:var(--amber-dark); }

    /* ── LESSON PREVIEW ── */
    .lesson-preview {
      background:white; border-radius:var(--radius);
      border:1px solid var(--border);
      padding:40px;
      min-height:500px;
      box-shadow: var(--shadow);
    }
    .lesson-title {
      font-family:'Playfair Display',serif;
      font-size:1.3rem; font-weight:700;
      text-align:center; text-transform:uppercase;
      margin-bottom:24px; color:var(--ink);
      line-height:1.4;
    }
    .lesson-section { margin-bottom:18px; }
    .lesson-section-head {
      font-size:0.95rem; font-weight:700;
      color:var(--amber-dark); margin-bottom:6px;
    }
    .lesson-body { font-size:0.9rem; line-height:1.8; color:var(--slate); }
    .lesson-body ul { padding-left:20px; }
    .lesson-body li { margin-bottom:4px; }
    .lesson-meta-grid {
      display:grid; grid-template-columns:repeat(3,1fr); gap:12px;
      background:var(--cream); border-radius:10px;
      padding:14px; margin-bottom:20px;
    }
    .meta-item label { font-size:0.72rem; font-weight:700; text-transform:uppercase; color:var(--muted); }
    .meta-item span { font-size:0.9rem; font-weight:600; display:block; margin-top:2px; }

    /* ── BUTTONS ── */
    .btn {
      padding:10px 20px; border-radius:10px; border:none;
      font-family:'DM Sans',sans-serif; font-size:0.9rem; font-weight:600;
      cursor:pointer; transition:all 0.2s;
    }
    .btn-amber {
      background: linear-gradient(135deg, var(--amber), var(--amber-dark));
      color:white; box-shadow: 0 2px 10px rgba(232,160,32,0.35);
    }
    .btn-amber:hover { transform:translateY(-1px); box-shadow: 0 4px 16px rgba(232,160,32,0.5); }
    .btn-amber:disabled { opacity:0.5; cursor:not-allowed; transform:none; }
    .btn-ghost { background:var(--cream); color:var(--ink); }
    .btn-ghost:hover { background:var(--border); }
    .btn-green { background:#059669; color:white; }
    .btn-green:hover { background:#047857; }

    /* ── ALERTS ── */
    .alert {
      padding:12px 16px; border-radius:10px;
      font-size:0.88rem; margin-bottom:16px;
      display:flex; align-items:center; gap:10px;
    }
    .alert-error { background:#fee2e2; color:#991b1b; border:1px solid #fca5a5; }
    .alert-success { background:#d1fae5; color:#065f46; border:1px solid #6ee7b7; }
    .alert-info { background:var(--amber-glow); color:var(--amber-dark); border:1px solid rgba(232,160,32,0.3); }

    /* ── SPINNER ── */
    .spinner {
      width:20px; height:20px; border-radius:50%;
      border:2px solid rgba(232,160,32,0.2);
      border-top-color: var(--amber);
      animation: spin 0.7s linear infinite;
      display:inline-block;
    }
    @keyframes spin { to { transform:rotate(360deg); } }

    /* ── TOPICS EDITOR ── */
    .topics-editor { background:white; border-radius:var(--radius); border:1px solid var(--border); overflow:hidden; }
    .topics-editor-head {
      padding:16px 20px; border-bottom:1px solid var(--border);
      display:flex; align-items:center; justify-content:space-between;
    }
    .topic-row {
      display:flex; align-items:center; gap:10px;
      padding:10px 20px; border-bottom:1px solid var(--border);
    }
    .topic-row:last-child { border-bottom:none; }
    .week-label {
      width:60px; font-size:0.78rem; font-weight:700;
      color:var(--amber-dark); flex-shrink:0;
    }
    .topic-input {
      flex:1; border:none; background:transparent;
      font-size:0.9rem; padding:6px 0;
      border-bottom:1.5px solid transparent;
    }
    .topic-input:focus { outline:none; border-bottom-color:var(--amber); }

    /* ── DOWNLOAD AREA ── */
    .download-bar {
      display:flex; align-items:center; justify-content:space-between;
      background:white; border-radius:var(--radius);
      border:1px solid var(--border); padding:16px 24px;
      margin-bottom:20px; box-shadow: var(--shadow);
    }
    .download-info { font-size:0.9rem; }
    .download-info strong { display:block; margin-bottom:2px; }
    .download-info span { color:var(--muted); font-size:0.82rem; }
    .download-actions { display:flex; gap:8px; }

    /* ── ADMIN ── */
    .admin-page { padding:40px; }
    .admin-table { width:100%; border-collapse:collapse; background:white; border-radius:var(--radius); overflow:hidden; box-shadow:var(--shadow); }
    .admin-table th {
      background:var(--ink); color:white;
      padding:12px 16px; text-align:left;
      font-size:0.82rem; font-weight:600; letter-spacing:0.05em;
    }
    .admin-table td { padding:12px 16px; border-bottom:1px solid var(--border); font-size:0.88rem; }
    .admin-table tr:last-child td { border-bottom:none; }
    .admin-table tr:hover td { background:var(--cream); }
    .stat-card {
      background:white; border-radius:var(--radius); border:1px solid var(--border);
      padding:24px; box-shadow:var(--shadow);
    }
    .stat-card .num { font-family:'Playfair Display',serif; font-size:2.4rem; font-weight:900; color:var(--amber-dark); }
    .stat-card .label { font-size:0.85rem; color:var(--muted); margin-top:4px; }
    .admin-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:20px; margin-bottom:32px; }

    /* ── STEP INDICATOR ── */
    .steps {
      display:flex; gap:0; margin-bottom:32px;
      background:white; border-radius:var(--radius);
      border:1px solid var(--border); overflow:hidden;
    }
    .step {
      flex:1; padding:14px; text-align:center;
      font-size:0.82rem; font-weight:600; color:var(--muted);
      border-right:1px solid var(--border);
      transition:all 0.2s;
    }
    .step:last-child { border-right:none; }
    .step.active { background:var(--amber-glow); color:var(--amber-dark); }
    .step.done { background:var(--cream); color:var(--ink); }
    .step-num {
      width:22px; height:22px; border-radius:50%;
      background:var(--cream); color:var(--muted);
      font-size:0.75rem; font-weight:700;
      display:inline-flex; align-items:center; justify-content:center;
      margin-right:8px;
    }
    .step.active .step-num { background:var(--amber); color:white; }
    .step.done .step-num { background:var(--forest); color:white; }

    /* ── TOAST ── */
    #toast {
      position:fixed; bottom:24px; right:24px; z-index:9999;
      background:var(--ink); color:white;
      padding:12px 20px; border-radius:10px;
      font-size:0.88rem; box-shadow:var(--shadow-lg);
      opacity:0; transform:translateY(10px);
      transition:all 0.3s; pointer-events:none;
    }
    #toast.show { opacity:1; transform:translateY(0); }

    /* ── MOBILE ── */
    @media(max-width:768px){
      nav { padding:0 20px; }
      .hero { grid-template-columns:1fr; padding:32px 24px; }
      .hero-visual { display:none; }
      .app-layout { grid-template-columns:1fr; }
      .sidebar { border-right:none; border-bottom:1px solid var(--border); }
      .form-grid { grid-template-columns:1fr; }
      .admin-grid { grid-template-columns:1fr 1fr; }
    }

    /* ── ANIMATIONS ── */
    @keyframes fadeIn { from{opacity:0;transform:translateY(12px)} to{opacity:1;transform:translateY(0)} }
    .fade-in { animation:fadeIn 0.4s ease forwards; }
    @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.5} }
    .pulse { animation:pulse 1.5s ease infinite; }
  </style>
</head>
<body>

<!-- NAVBAR -->
<nav>
  <div class="logo">Teachly</div>
  <div class="nav-links">
    <button class="nav-btn ghost" onclick="showPage('home')">Home</button>
    <button class="nav-btn ghost" onclick="showPage('app')">Dashboard</button>
    <button class="nav-btn primary" onclick="showPage('app')">Get Started →</button>
  </div>
</nav>

<!-- TOAST -->
<div id="toast"></div>

<!-- ══════════════════════════════════════════════════════ -->
<!-- HOME PAGE -->
<!-- ══════════════════════════════════════════════════════ -->
<div id="page-home" class="page active">
  <div class="hero">
    <div class="hero-text fade-in">
      <div class="hero-badge">Built for Nigerian Teachers</div>
      <h1>Generate <span>Lesson Notes</span> in Minutes, Not Hours</h1>
      <p class="hero-sub">Upload your Scheme of Work and let Teachly's AI write complete, inspection-ready lesson notes for every week of the term.</p>
      <div class="hero-cta">
        <button class="btn-lg btn-primary" onclick="showPage('app')">Start Generating Free →</button>
        <button class="btn-lg btn-outline" onclick="showDemo()">See a Sample Note</button>
      </div>
      <div class="stats">
        <div><div class="stat-num">40+</div><div class="stat-label">Weeks per term covered</div></div>
        <div><div class="stat-num">10×</div><div class="stat-label">Faster than manual writing</div></div>
        <div><div class="stat-num">100%</div><div class="stat-label">Nigerian curriculum aligned</div></div>
      </div>
    </div>
    <div class="hero-visual fade-in">
      <div class="mock-header">
        <div class="mock-dot" style="background:#ff5f57"></div>
        <div class="mock-dot" style="background:#febc2e"></div>
        <div class="mock-dot" style="background:#28c840"></div>
        <span style="margin-left:8px;font-size:0.8rem;color:var(--muted)">Teachly — SS2 Biology</span>
      </div>
      <div class="mock-week-list">
        <div class="mock-week active">
          <div class="mock-week-num">1</div>
          <span>Cell Structure & Organisation</span>
          <span class="mock-status done">✓ Done</span>
        </div>
        <div class="mock-week active">
          <div class="mock-week-num">2</div>
          <span>Cell Division – Mitosis</span>
          <span class="mock-status done">✓ Done</span>
        </div>
        <div class="mock-week">
          <div class="mock-week-num" style="background:var(--cream);color:var(--muted)">3</div>
          <span>Cell Division – Meiosis</span>
          <span class="mock-status gen pulse">⚡ Generating…</span>
        </div>
        <div class="mock-week">
          <div class="mock-week-num" style="background:var(--cream);color:var(--muted)">4</div>
          <span>Nutrition in Living Things</span>
          <span class="mock-status" style="color:var(--muted)">Queued</span>
        </div>
        <div class="mock-week">
          <div class="mock-week-num" style="background:var(--cream);color:var(--muted)">5</div>
          <span>Photosynthesis</span>
          <span class="mock-status" style="color:var(--muted)">Queued</span>
        </div>
      </div>
      <div style="margin-top:20px;padding:16px;background:white;border-radius:10px;border:1px solid var(--border);">
        <div style="font-size:0.72rem;font-weight:700;text-transform:uppercase;color:var(--muted);margin-bottom:8px;">Step 2: Exploration</div>
        <div style="font-size:0.82rem;line-height:1.7;color:var(--slate)">
          Cell division is the process by which a parent cell divides into two daughter cells. Mitosis produces genetically identical cells and is used for growth and repair in Nigerian plants and animals alike…
        </div>
      </div>
    </div>
  </div>
</div>


<!-- ══════════════════════════════════════════════════════ -->
<!-- APP PAGE -->
<!-- ══════════════════════════════════════════════════════ -->
<div id="page-app" class="page">
  <div class="app-layout">
    <!-- SIDEBAR -->
    <div class="sidebar" id="sidebar">
      <div class="sidebar-title">Weeks</div>
      <div id="week-list">
        <div style="color:var(--muted);font-size:0.85rem;padding:20px 0;text-align:center">
          Upload a scheme of work to get started
        </div>
      </div>
      <div style="margin-top:20px;border-top:1px solid var(--border);padding-top:20px;" id="download-section" style="display:none">
        <div class="sidebar-title">Export</div>
        <button class="btn btn-green" style="width:100%;margin-bottom:8px" onclick="downloadAllDocx()">
          ⬇ Download Full Term (DOCX)
        </button>
        <button class="btn btn-ghost" style="width:100%" onclick="downloadCurrentDocx()">
          ⬇ Download This Week
        </button>
      </div>
    </div>

    <!-- MAIN -->
    <div class="main-content">
      <div class="steps" id="steps-indicator">
        <div class="step active" id="step1"><span class="step-num">1</span>Upload</div>
        <div class="step" id="step2"><span class="step-num">2</span>Review Topics</div>
        <div class="step" id="step3"><span class="step-num">3</span>Details</div>
        <div class="step" id="step4"><span class="step-num">4</span>Generate</div>
      </div>

      <!-- STEP 1: UPLOAD -->
      <div id="section-upload">
        <div id="upload-alert"></div>
        <div class="upload-zone" id="upload-zone" onclick="document.getElementById('file-input').click()" ondragover="handleDragOver(event)" ondragleave="handleDragLeave(event)" ondrop="handleDrop(event)">
          <div class="upload-icon">📄</div>
          <div class="upload-title">Drop your Scheme of Work here</div>
          <div class="upload-sub">or click to browse your files</div>
          <div class="upload-types">
            <span class="type-badge">PDF</span>
            <span class="type-badge">DOCX</span>
            <span class="type-badge">PNG</span>
            <span class="type-badge">JPEG</span>
          </div>
        </div>
        <input type="file" id="file-input" style="display:none" accept=".pdf,.docx,.doc,.png,.jpg,.jpeg" onchange="handleFileSelect(event)">
        
        <div id="upload-progress" style="display:none;margin-top:16px">
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">
            <div class="spinner"></div>
            <span id="upload-status" style="font-size:0.9rem">Processing file…</span>
          </div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-fill" id="progress-fill" style="width:0%"></div>
          </div>
        </div>
      </div>

      <!-- STEP 2: TOPICS EDITOR -->
      <div id="section-topics" style="display:none">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
          <div>
            <h2 style="font-family:'Playfair Display',serif;font-size:1.4rem">Review & Edit Topics</h2>
            <p style="font-size:0.85rem;color:var(--muted);margin-top:4px">Verify extracted topics. Edit any that are incorrect.</p>
          </div>
          <div style="display:flex;gap:8px">
            <button class="btn btn-ghost" onclick="backToUpload()">← Back</button>
            <button class="btn btn-amber" onclick="proceedToDetails()">Continue →</button>
          </div>
        </div>
        <div class="topics-editor" id="topics-container"></div>
      </div>

      <!-- STEP 3: DETAILS FORM -->
      <div id="section-details" style="display:none">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
          <div>
            <h2 style="font-family:'Playfair Display',serif;font-size:1.4rem">Lesson Details</h2>
            <p style="font-size:0.85rem;color:var(--muted);margin-top:4px">Tell Teachly about your class so it can personalise the notes.</p>
          </div>
          <button class="btn btn-ghost" onclick="backToTopics()">← Back</button>
        </div>
        <div class="form-grid">
          <div class="form-group">
            <label>Subject *</label>
            <input type="text" id="input-subject" placeholder="e.g. Biology, Mathematics, English">
          </div>
          <div class="form-group">
            <label>Class *</label>
            <select id="input-class">
              <option value="">Select class…</option>
              <optgroup label="Primary">
                <option>Primary 1</option><option>Primary 2</option><option>Primary 3</option>
                <option>Primary 4</option><option>Primary 5</option><option>Primary 6</option>
              </optgroup>
              <optgroup label="Junior Secondary">
                <option>JSS 1</option><option>JSS 2</option><option>JSS 3</option>
              </optgroup>
              <optgroup label="Senior Secondary">
                <option>SS 1</option><option>SS 2</option><option>SS 3</option>
              </optgroup>
            </select>
          </div>
          <div class="form-group">
            <label>Term *</label>
            <select id="input-term">
              <option value="">Select term…</option>
              <option>1st Term</option><option>2nd Term</option><option>3rd Term</option>
            </select>
          </div>
          <div class="form-group">
            <label>School Name (optional)</label>
            <input type="text" id="input-school" placeholder="e.g. WOF Schools">
          </div>
          <div class="form-group">
            <label>Teacher's Name (optional)</label>
            <input type="text" id="input-teacher" placeholder="e.g. Mr. Chinedu">
          </div>
          <div class="form-group">
            <label>Term Start Date</label>
            <input type="date" id="input-start-date">
          </div>
        </div>
        <div style="margin-top:24px">
          <button class="btn btn-amber btn-lg" id="generate-btn" onclick="startGeneration()">
            ⚡ Generate All Lesson Notes
          </button>
        </div>
      </div>

      <!-- STEP 4: PREVIEW -->
      <div id="section-preview" style="display:none">
        <div id="generation-progress" style="margin-bottom:24px;display:none">
          <div class="alert alert-info">
            <div class="spinner"></div>
            <div>
              <strong id="gen-status-text">Generating lesson notes…</strong>
              <div id="gen-sub-text" style="font-size:0.8rem;margin-top:2px;opacity:0.8">Please wait while AI writes your notes</div>
            </div>
          </div>
          <div class="progress-bar-wrap">
            <div class="progress-bar-fill" id="gen-progress-fill" style="width:0%"></div>
          </div>
        </div>
        
        <div id="download-bar-wrap" style="display:none">
          <div class="download-bar">
            <div class="download-info">
              <strong>🎉 All lesson notes generated!</strong>
              <span id="dl-summary">Ready to download</span>
            </div>
            <div class="download-actions">
              <button class="btn btn-ghost" onclick="downloadCurrentDocx()">⬇ This Week</button>
              <button class="btn btn-green" onclick="downloadAllDocx()">⬇ Full Term DOCX</button>
            </div>
          </div>
        </div>

        <div id="lesson-display" class="lesson-preview">
          <div style="text-align:center;color:var(--muted);padding:60px 0">
            Select a week from the sidebar to preview its lesson note
          </div>
        </div>
      </div>

    </div>
  </div>
</div>


<!-- ══════════════════════════════════════════════════════ -->
<!-- ADMIN PAGE -->
<!-- ══════════════════════════════════════════════════════ -->
<div id="page-admin" class="page">
  <div class="admin-page">
    <div style="margin-bottom:32px">
      <h1 style="font-family:'Playfair Display',serif;font-size:2rem;margin-bottom:6px">Admin Dashboard</h1>
      <p style="color:var(--muted)">Teachly platform overview</p>
    </div>
    <div class="admin-grid">
      <div class="stat-card"><div class="num" id="admin-users">0</div><div class="label">Total Users</div></div>
      <div class="stat-card"><div class="num" id="admin-uploads">0</div><div class="label">Total Uploads</div></div>
      <div class="stat-card"><div class="num" id="admin-generated">0</div><div class="label">Notes Generated</div></div>
      <div class="stat-card"><div class="num" id="admin-downloads">0</div><div class="label">Downloads</div></div>
    </div>
    <h3 style="margin-bottom:16px;font-size:1rem;font-weight:700">Recent Activity</h3>
    <table class="admin-table">
      <thead>
        <tr>
          <th>Session</th><th>Subject</th><th>Class</th><th>Weeks</th><th>Notes Generated</th><th>Downloads</th><th>Time</th>
        </tr>
      </thead>
      <tbody id="admin-table-body">
        <tr><td colspan="7" style="text-align:center;color:var(--muted);padding:32px">No activity yet</td></tr>
      </tbody>
    </table>
  </div>
</div>


<script>
// ═══════════════════════════════════════════════════════
//  STATE
// ═══════════════════════════════════════════════════════
const state = {
  weeks: [],           // [{week, topic}]
  lessonNotes: {},     // {weekNum: "generated text"}
  currentWeek: null,
  subject: '', className: '', term: '', school: '', teacher: '', startDate: '',
  step: 1,
  stats: { uploads:0, generated:0, downloads:0 },
  activity: [],
  apiKey: null
};

// ═══════════════════════════════════════════════════════
//  NAVIGATION
// ═══════════════════════════════════════════════════════
function showPage(name) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  if(name==='admin') loadAdminData();
}

// Hidden admin route via URL hash
window.addEventListener('hashchange', () => {
  if(location.hash === '#/admin') showPage('admin');
});
if(location.hash === '#/admin') showPage('admin');

function setStep(n) {
  state.step = n;
  ['1','2','3','4'].forEach(i => {
    const el = document.getElementById('step'+i);
    el.classList.remove('active','done');
    if(+i < n) el.classList.add('done');
    else if(+i === n) el.classList.add('active');
  });
  document.getElementById('section-upload').style.display = n===1?'block':'none';
  document.getElementById('section-topics').style.display = n===2?'block':'none';
  document.getElementById('section-details').style.display = n===3?'block':'none';
  document.getElementById('section-preview').style.display = n===4?'block':'none';
}

function backToUpload() { setStep(1); }
function backToTopics() { setStep(2); }

// ═══════════════════════════════════════════════════════
//  FILE UPLOAD
// ═══════════════════════════════════════════════════════
function handleDragOver(e) { e.preventDefault(); document.getElementById('upload-zone').classList.add('drag-over'); }
function handleDragLeave() { document.getElementById('upload-zone').classList.remove('drag-over'); }
function handleDrop(e) {
  e.preventDefault();
  document.getElementById('upload-zone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if(file) processFile(file);
}
function handleFileSelect(e) {
  const file = e.target.files[0];
  if(file) processFile(file);
}

async function processFile(file) {
  const allowed = ['application/pdf','application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/msword','image/png','image/jpeg','image/jpg'];
  if(!allowed.includes(file.type)) {
    showAlert('upload-alert','error','❌ Unsupported file type. Please upload PDF, DOCX, PNG, or JPEG.');
    return;
  }
  state.stats.uploads++;
  showUploadProgress(true, 'Reading file…', 20);
  clearAlert('upload-alert');
  
  try {
    let text = '';
    if(file.type.includes('pdf')) {
      text = await extractFromPDF(file);
    } else if(file.type.includes('word') || file.name.endsWith('.docx') || file.name.endsWith('.doc')) {
      text = await extractFromDOCX(file);
    } else {
      text = await extractFromImage(file);
    }
    
    showUploadProgress(true, 'Parsing topics…', 70);
    
    if(!text || text.trim().length < 20) {
      // Fallback: ask user to type topics
      showAlert('upload-alert','info','⚠ Could not auto-extract text. You can manually enter topics in the next step.');
      state.weeks = generateEmptyWeeks(13);
    } else {
      state.weeks = parseTopicsFromText(text);
    }
    
    showUploadProgress(true, 'Done!', 100);
    setTimeout(() => {
      showUploadProgress(false);
      renderTopicsEditor();
      setStep(2);
    }, 500);
    
  } catch(err) {
    showUploadProgress(false);
    showAlert('upload-alert','error','❌ Failed to process file: ' + err.message);
  }
}

async function extractFromDOCX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const result = await mammoth.extractRawText({arrayBuffer: e.target.result});
        resolve(result.value);
      } catch(err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

async function extractFromPDF(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        // Simple text extraction attempt using ArrayBuffer
        const bytes = new Uint8Array(e.target.result);
        let text = '';
        // Extract readable ASCII text from PDF bytes
        for(let i=0; i<bytes.length-1; i++) {
          const c = bytes[i];
          if(c >= 32 && c < 127) text += String.fromCharCode(c);
          else if(c === 10 || c === 13) text += ' ';
        }
        // Try to pull out recognizable content
        const cleaned = text.replace(/[^\x20-\x7E\n]/g,' ').replace(/\s+/g,' ');
        resolve(cleaned);
      } catch(err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

async function extractFromImage(file) {
  // Since we can't run Tesseract without a CDN in this environment,
  // we'll use the Anthropic vision API to extract text from the image
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const base64 = e.target.result.split(',')[1];
        const mtype = file.type;
        const response = await fetch('https://api.anthropic.com/v1/messages', {
          method:'POST',
          headers: { 'Content-Type':'application/json' },
          body: JSON.stringify({
            model: 'claude-sonnet-4-20250514',
            max_tokens: 2000,
            messages: [{
              role:'user',
              content: [
                { type:'image', source:{ type:'base64', media_type:mtype, data:base64 } },
                { type:'text', text:'Extract ALL text from this scheme of work image. Output the raw text only, preserving week numbers and topic names.' }
              ]
            }]
          })
        });
        const data = await response.json();
        resolve(data.content?.[0]?.text || '');
      } catch(err) { reject(err); }
    };
    reader.readAsDataURL(file);
  });
}

function parseTopicsFromText(text) {
  const weeks = [];
  // Match patterns like: Week 1: Topic, WEEK 1 - Topic, 1. Topic, etc.
  const patterns = [
    /week\s*(\d+)[:\-–\s]+([^\n\r]+)/gi,
    /wk\.?\s*(\d+)[:\-–\s]+([^\n\r]+)/gi,
    /(\d+)\.\s+([A-Z][^\n\r]{5,60})/gm
  ];
  
  for(const pattern of patterns) {
    let match;
    while((match = pattern.exec(text)) !== null) {
      const wn = parseInt(match[1]);
      const topic = match[2].trim().replace(/[^\w\s\-&,()]/g,'').trim();
      if(topic.length > 3 && wn >= 1 && wn <= 40) {
        if(!weeks.find(w => w.week === wn)) {
          weeks.push({ week: wn, topic });
        }
      }
    }
    if(weeks.length >= 3) break;
  }
  
  weeks.sort((a,b) => a.week - b.week);
  
  // If we got at least 3 weeks, return them; otherwise return empty
  if(weeks.length >= 3) return weeks;
  
  // Try line-based fallback
  const lines = text.split(/[\n\r]+/).map(l=>l.trim()).filter(l=>l.length>5);
  const fallback = [];
  let wkNum = 1;
  for(const line of lines) {
    if(line.length > 4 && line.length < 120 && !/^[0-9\s\-\.]+$/.test(line)) {
      fallback.push({ week: wkNum, topic: line.substring(0,80) });
      wkNum++;
      if(wkNum > 40) break;
    }
  }
  return fallback.length >= 3 ? fallback.slice(0,40) : generateEmptyWeeks(13);
}

function generateEmptyWeeks(n) {
  return Array.from({length:n}, (_,i) => ({ week:i+1, topic:'' }));
}

// ═══════════════════════════════════════════════════════
//  TOPICS EDITOR
// ═══════════════════════════════════════════════════════
function renderTopicsEditor() {
  const c = document.getElementById('topics-container');
  c.innerHTML = `
    <div class="topics-editor-head">
      <span style="font-weight:600">📋 ${state.weeks.length} weeks detected</span>
      <button class="btn btn-ghost" style="font-size:0.82rem;padding:6px 12px" onclick="addWeek()">+ Add Week</button>
    </div>
    ${state.weeks.map((w,i) => `
      <div class="topic-row">
        <span class="week-label">Week ${w.week}</span>
        <input class="topic-input" id="topic-${i}" value="${escHtml(w.topic)}" placeholder="Enter topic for week ${w.week}…">
        <button onclick="removeWeek(${i})" style="background:none;border:none;cursor:pointer;color:var(--muted);font-size:1.1rem;padding:4px" title="Remove">✕</button>
      </div>
    `).join('')}
  `;
}

function addWeek() {
  const lastWk = state.weeks.length ? state.weeks[state.weeks.length-1].week + 1 : 1;
  state.weeks.push({ week:lastWk, topic:'' });
  renderTopicsEditor();
}
function removeWeek(i) {
  state.weeks.splice(i,1);
  renderTopicsEditor();
}

function proceedToDetails() {
  // Save edits
  state.weeks = state.weeks.map((w,i) => ({
    week: w.week,
    topic: (document.getElementById('topic-'+i)?.value || w.topic).trim()
  })).filter(w => w.topic.length > 0);
  
  if(state.weeks.length === 0) {
    showAlert('upload-alert','error','Please enter at least one topic before continuing.');
    setStep(2); return;
  }
  setStep(3);
}

// ═══════════════════════════════════════════════════════
//  GENERATION
// ═══════════════════════════════════════════════════════
async function startGeneration() {
  const subject = document.getElementById('input-subject').value.trim();
  const className = document.getElementById('input-class').value;
  const term = document.getElementById('input-term').value;
  
  if(!subject || !className || !term) {
    showToast('Please fill in Subject, Class, and Term.');
    return;
  }
  
  state.subject = subject;
  state.className = className;
  state.term = term;
  state.school = document.getElementById('input-school').value.trim();
  state.teacher = document.getElementById('input-teacher').value.trim();
  state.startDate = document.getElementById('input-start-date').value;
  
  setStep(4);
  renderWeekList();
  
  document.getElementById('generation-progress').style.display = 'block';
  document.getElementById('download-bar-wrap').style.display = 'none';
  
  const total = state.weeks.length;
  let done = 0;
  
  for(const wk of state.weeks) {
    document.getElementById('gen-status-text').textContent = `Generating Week ${wk.week}: ${wk.topic}…`;
    document.getElementById('gen-sub-text').textContent = `${done}/${total} complete`;
    
    try {
      const note = await generateLessonNote(wk, subject, className, term);
      state.lessonNotes[wk.week] = note;
      state.stats.generated++;
      done++;
      const pct = Math.round((done/total)*100);
      document.getElementById('gen-progress-fill').style.width = pct+'%';
      updateWeekStatus(wk.week, 'done');
      if(state.currentWeek === null) {
        state.currentWeek = wk.week;
        renderLesson(wk.week);
      }
    } catch(err) {
      state.lessonNotes[wk.week] = `[Error generating note for Week ${wk.week}: ${err.message}]`;
      done++;
      updateWeekStatus(wk.week, 'error');
    }
    renderWeekList();
  }
  
  document.getElementById('generation-progress').style.display = 'none';
  document.getElementById('download-bar-wrap').style.display = 'block';
  document.getElementById('dl-summary').textContent = `${total} lesson notes ready for ${subject} – ${className}`;
  document.getElementById('download-section').style.display = 'block';
  
  // Log activity
  logActivity(subject, className, total, done);
  showToast('✅ All lesson notes generated!');
}

async function generateLessonNote(wk, subject, className, term) {
  const startDate = state.startDate ? new Date(state.startDate) : new Date();
  const weekDate = new Date(startDate);
  weekDate.setDate(weekDate.getDate() + (wk.week - 1) * 7 + 4); // Friday of that week
  const dateStr = weekDate.toLocaleDateString('en-GB', { day:'2-digit', month:'long', year:'numeric' });
  
  const prompt = `You are an experienced Nigerian secondary/primary school teacher writing a formal lesson note.

Write a COMPLETE, DETAILED lesson note following this EXACT format:

---
NOTE OF LESSON FOR ${subject.toUpperCase()} FOR WEEK ${wk.week} ENDING ${dateStr}

Class: ${className}
Date: ${dateStr}
Topic: ${wk.topic}
Sub-topic: [relevant sub-topic]
Gender: Mixed
Duration: 40 minutes
---

1. SPECIFIC OBJECTIVES
By the end of the lesson, students should be able to:
i. [objective 1]
ii. [objective 2]
iii. [objective 3]
iv. [objective 4]

2. BIBLE REFERENCE
[Relevant Bible verse that connects to the lesson theme]

3. INSTRUCTIONAL MATERIALS
[List at least 5 relevant materials/resources]

4. INSTRUCTIONAL TECHNIQUES
[List techniques: lecture, demonstration, discussion, question-and-answer, etc.]

5. INSTRUCTIONAL PROCEDURES

Step 1: Identification of Prior Idea
[3-4 classroom questions to assess prior knowledge]

Step 2: Exploration
[VERY DETAILED explanation - minimum 300 words. Define ALL key terms. Use Nigerian examples. Break down every concept clearly. Explain as a master teacher would.]

Step 3: Class Activity & Discussion
[Practical activity or group discussion exercise]

Step 4: Application
[How this topic applies in Nigerian daily life or environment]

Step 5: Evaluation
[5 assessment questions - mix of objective and theory]

STUDENTS' ROLE
[What students should do during the lesson]

SUMMARY
[Brief summary of the lesson]

ASSIGNMENT
[2-3 homework questions]

APPENDIX
[Additional notes, diagrams descriptions, or reference materials]

---

Topic: ${wk.topic}
Subject: ${subject}
Class: ${className}
Term: ${term}
School: ${state.school || 'Secondary School'}
Teacher: ${state.teacher || 'Subject Teacher'}

Write the FULL, COMPLETE lesson note now. Be thorough, detailed, and professional. Use Nigerian curriculum context throughout.`;

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [{ role: 'user', content: prompt }]
    })
  });
  
  if(!response.ok) throw new Error(`API error: ${response.status}`);
  const data = await response.json();
  return data.content?.[0]?.text || 'Error: no content returned';
}

// ═══════════════════════════════════════════════════════
//  WEEK LIST & PREVIEW
// ═══════════════════════════════════════════════════════
function renderWeekList() {
  const container = document.getElementById('week-list');
  container.innerHTML = state.weeks.map(w => {
    const hasNote = !!state.lessonNotes[w.week];
    const isActive = state.currentWeek === w.week;
    let statusClass = '', statusText = 'Pending';
    if(hasNote) { statusClass='done'; statusText='✓ Ready'; }
    else if(state.step===4) { statusClass='gen'; statusText='⚡ Queued'; }
    
    return `
      <div class="week-item ${isActive?'active':''}" onclick="selectWeek(${w.week})">
        <div class="week-badge">${w.week}</div>
        <div class="week-info">
          <div class="week-name">${w.topic || 'Week '+w.week}</div>
          <div class="week-status ${statusClass}">${statusText}</div>
        </div>
      </div>
    `;
  }).join('');
}

function updateWeekStatus(weekNum, status) {
  // Will be re-rendered via renderWeekList()
}

function selectWeek(wn) {
  state.currentWeek = wn;
  renderWeekList();
  if(state.lessonNotes[wn]) renderLesson(wn);
  else {
    document.getElementById('lesson-display').innerHTML = `
      <div style="text-align:center;color:var(--muted);padding:60px 0">
        <div class="spinner" style="margin:0 auto 16px"></div>
        <div>Week ${wn} is being generated…</div>
      </div>
    `;
  }
}

function renderLesson(wn) {
  const note = state.lessonNotes[wn];
  if(!note) return;
  const wk = state.weeks.find(w=>w.week===wn);
  const display = document.getElementById('lesson-display');
  
  // Format the note as HTML
  const lines = note.split('\n');
  let html = '<div class="fade-in">';
  let inList = false;
  
  for(const line of lines) {
    const trimmed = line.trim();
    if(!trimmed) {
      if(inList) { html += '</ul>'; inList=false; }
      html += '<div style="height:8px"></div>';
      continue;
    }
    
    if(trimmed.startsWith('NOTE OF LESSON')) {
      if(inList) { html += '</ul>'; inList=false; }
      html += `<div class="lesson-title">${escHtml(trimmed)}</div>`;
    } else if(/^\d+\.\s+[A-Z]/.test(trimmed) || /^Step \d+:/.test(trimmed)) {
      if(inList) { html += '</ul>'; inList=false; }
      html += `<div class="lesson-section-head">${escHtml(trimmed)}</div>`;
    } else if(trimmed.startsWith('---')) {
      if(inList) { html += '</ul>'; inList=false; }
      html += '<hr style="border:none;border-top:1px solid var(--border);margin:16px 0">';
    } else if(/^[ivxlIVX]+\./.test(trimmed) || trimmed.startsWith('- ') || trimmed.startsWith('• ')) {
      if(!inList) { html += '<ul class="lesson-body" style="padding-left:24px">'; inList=true; }
      html += `<li>${escHtml(trimmed.replace(/^[-•ivxlIVX\d]+[.)]\s*/,''))}</li>`;
    } else {
      if(inList) { html += '</ul>'; inList=false; }
      html += `<p class="lesson-body" style="margin-bottom:6px">${escHtml(trimmed)}</p>`;
    }
  }
  if(inList) html += '</ul>';
  html += '</div>';
  
  display.innerHTML = html;
}

// ═══════════════════════════════════════════════════════
//  DOCX EXPORT
// ═══════════════════════════════════════════════════════
function downloadCurrentDocx() {
  if(!state.currentWeek || !state.lessonNotes[state.currentWeek]) {
    showToast('No lesson note selected'); return;
  }
  const wk = state.weeks.find(w=>w.week===state.currentWeek);
  buildAndDownloadDocx([{week:wk, note:state.lessonNotes[state.currentWeek]}],
    `Teachly_${state.subject}_Week${state.currentWeek}.docx`);
  state.stats.downloads++;
}

function downloadAllDocx() {
  const ready = Object.keys(state.lessonNotes).map(wn => ({
    week: state.weeks.find(w=>w.week===+wn),
    note: state.lessonNotes[wn]
  })).filter(x=>x.week);
  
  if(!ready.length) { showToast('No notes to download yet'); return; }
  buildAndDownloadDocx(ready, `Teachly_${state.subject}_${state.className}_${state.term}.docx`);
  state.stats.downloads++;
}

function buildAndDownloadDocx(items, filename) {
  try {
    const { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, Packer, PageBreak } = docx;
    
    const children = [];
    
    for(const item of items) {
      const lines = item.note.split('\n');
      
      for(const line of lines) {
        const trimmed = line.trim();
        if(!trimmed) { children.push(new Paragraph({ text:'' })); continue; }
        
        if(trimmed.startsWith('NOTE OF LESSON')) {
          children.push(new Paragraph({
            children:[new TextRun({text:trimmed, bold:true, size:36, font:'Times New Roman', allCaps:true})],
            alignment: AlignmentType.CENTER, spacing:{after:200}
          }));
        } else if(/^\d+\.\s+[A-Z]/.test(trimmed) || /^Step \d+:/.test(trimmed) ||
                  ['STUDENTS\' ROLE','SUMMARY','ASSIGNMENT','APPENDIX'].some(h=>trimmed.startsWith(h))) {
          children.push(new Paragraph({
            children:[new TextRun({text:trimmed, bold:true, size:32, font:'Times New Roman'})],
            spacing:{before:200, after:100}
          }));
        } else if(trimmed.startsWith('---')) {
          children.push(new Paragraph({text:'', border:{bottom:{value:'single',size:6,color:'CCCCCC'}}}));
        } else {
          children.push(new Paragraph({
            children:[new TextRun({text:trimmed, size:26, font:'Times New Roman'})],
            alignment: AlignmentType.JUSTIFY,
            spacing:{after:80}
          }));
        }
      }
      
      // Page break between weeks
      if(items.indexOf(item) < items.length-1) {
        children.push(new Paragraph({children:[new PageBreak()]}));
      }
    }
    
    const doc = new Document({
      sections:[{
        properties:{
          page:{
            margin:{ top:1440, bottom:1440, left:1440, right:1440 }
          }
        },
        children
      }]
    });
    
    Packer.toBlob(doc).then(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast('📥 Download started!');
    });
    
  } catch(err) {
    // Fallback: plain text download
    const content = items.map(i=>i.note).join('\n\n' + '='.repeat(60) + '\n\n');
    const blob = new Blob([content], {type:'text/plain'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename.replace('.docx','.txt');
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('📥 Downloaded as text file');
  }
}

// ═══════════════════════════════════════════════════════
//  ADMIN
// ═══════════════════════════════════════════════════════
function logActivity(subject, className, weeks, generated) {
  const record = {
    id: Date.now(),
    subject, className, weeks, generated,
    downloads: 0,
    time: new Date().toLocaleString('en-NG')
  };
  state.activity.unshift(record);
  
  // Persist in localStorage
  try {
    const saved = JSON.parse(localStorage.getItem('teachly_activity')||'[]');
    saved.unshift(record);
    localStorage.setItem('teachly_activity', JSON.stringify(saved.slice(0,100)));
    const stats = JSON.parse(localStorage.getItem('teachly_stats')||'{}');
    stats.uploads = (stats.uploads||0) + 1;
    stats.generated = (stats.generated||0) + generated;
    localStorage.setItem('teachly_stats', JSON.stringify(stats));
  } catch(e) {}
}

function loadAdminData() {
  try {
    const activity = JSON.parse(localStorage.getItem('teachly_activity')||'[]');
    const stats = JSON.parse(localStorage.getItem('teachly_stats')||'{}');
    
    document.getElementById('admin-users').textContent = 1;
    document.getElementById('admin-uploads').textContent = stats.uploads || 0;
    document.getElementById('admin-generated').textContent = stats.generated || 0;
    document.getElementById('admin-downloads').textContent = stats.downloads || 0;
    
    const tbody = document.getElementById('admin-table-body');
    if(!activity.length) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:32px">No activity yet</td></tr>';
      return;
    }
    tbody.innerHTML = activity.map(a => `
      <tr>
        <td>#${a.id.toString().slice(-6)}</td>
        <td>${escHtml(a.subject)}</td>
        <td>${escHtml(a.className)}</td>
        <td>${a.weeks}</td>
        <td>${a.generated}</td>
        <td>${a.downloads}</td>
        <td>${a.time}</td>
      </tr>
    `).join('');
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════
//  DEMO
// ═══════════════════════════════════════════════════════
function showDemo() {
  showPage('app');
  setStep(4);
  state.weeks = [{week:1, topic:'Cell Structure and Organisation'}];
  state.currentWeek = 1;
  state.subject = 'Biology';
  state.className = 'SS 2';
  state.term = '1st Term';
  
  state.lessonNotes[1] = `NOTE OF LESSON FOR BIOLOGY FOR WEEK 1 ENDING 14TH JANUARY, 2025

Class: SS 2
Date: 14th January, 2025
Topic: Cell Structure and Organisation
Sub-topic: Structure and Functions of Cell Organelles
Gender: Mixed
Duration: 40 minutes

---

1. SPECIFIC OBJECTIVES
By the end of the lesson, students should be able to:
i. Define a cell and explain its importance as the basic unit of life
ii. Identify and describe the major organelles found in plant and animal cells
iii. Explain the functions of each cell organelle with reference to the human body
iv. Compare and contrast plant cells and animal cells

2. BIBLE REFERENCE
"For we are God's handiwork, created in Christ Jesus to do good works." – Ephesians 2:10
Just as God carefully designed each human being, He also designed each cell with specific organelles that work together for a purpose.

3. INSTRUCTIONAL MATERIALS
- Biology textbook (Senior Secondary School Biology, Book 2)
- Wall chart showing animal and plant cell diagrams
- Microscope slides of onion cells and cheek cells
- Handout with labelled cell diagrams
- Markers and whiteboard

4. INSTRUCTIONAL TECHNIQUES
- Lecture method
- Demonstration with microscope
- Question and answer
- Class discussion
- Use of diagrams and charts

5. INSTRUCTIONAL PROCEDURES

Step 1: Identification of Prior Idea
The teacher will ask students the following questions to assess prior knowledge:
- What do you think the smallest unit of a living organism is?
- Have you ever looked at anything under a microscope?
- Name one part of your body and describe what you think makes it up
- What do you think makes humans different from rocks and stones?

Step 2: Exploration
A CELL is defined as the smallest structural and functional unit of all living organisms. Robert Hooke first discovered cells in 1665 when he observed cork under a microscope and noticed small box-like compartments which he called cells. Later, Matthias Schleiden and Theodor Schwann developed the Cell Theory which states that: (1) all living things are made of cells, (2) the cell is the basic unit of life, and (3) all cells come from pre-existing cells.

KEY TERMS:
- Organelle: A specialised structure within a cell that performs a specific function, just like organs in the human body perform specific functions.
- Nucleus: The control centre of the cell. It contains DNA (genetic material) and controls all cell activities. Think of it as the head teacher of WOF Schools who gives instructions to everyone.
- Cell Membrane: A thin, flexible boundary that surrounds the cell. It controls what enters and leaves the cell. It is selectively permeable.
- Cell Wall: Found only in plant cells. It is made of cellulose and provides rigidity and support. This is why plant stems stand upright.
- Cytoplasm: The jelly-like fluid that fills the cell. All organelles float in the cytoplasm.
- Mitochondria: The powerhouse of the cell. It produces energy (ATP) through cellular respiration. Very active cells like muscle cells have many mitochondria.
- Chloroplast: Found only in plant cells. Contains chlorophyll which absorbs sunlight for photosynthesis. This is why leaves are green.
- Vacuole: A storage organelle. Plant cells have a large central vacuole for water storage. Animal cells have smaller vacuoles.
- Ribosome: Site of protein synthesis. Every cell needs proteins to function and grow.
- Endoplasmic Reticulum: A network of membranes that transports materials within the cell.

In Nigeria, a good analogy for a cell is a local government office: the nucleus is the LG chairman, the cell membrane is the security gate, different departments are like organelles, and the compound (cytoplasm) holds everything together.

Step 3: Class Activity & Discussion
Students will work in pairs to:
- Draw and label a plant cell and an animal cell
- Complete a comparison table: list 5 differences between plant and animal cells
- Observe prepared slides under the microscope (if available)

Step 4: Application
Cells are relevant to everyday Nigerian life in many ways:
- Farmers need to understand plant cells to know why fertilisers improve crop growth
- Understanding cell structure helps explain why eating protein-rich foods like beans and eggs helps the body repair tissues
- Medical laboratory scientists in Nigerian hospitals examine blood cells to diagnose diseases like sickle cell anaemia, which is very common in Nigeria

Step 5: Evaluation
1. Define a cell and state who first discovered it.
2. List FIVE organelles found in both plant and animal cells.
3. State TWO differences between a plant cell and an animal cell.
4. What is the function of the mitochondria?
5. Explain why plant cells have chloroplasts but animal cells do not.

STUDENTS' ROLE
Students will listen attentively, take notes, answer questions, participate in class discussions, draw diagrams of plant and animal cells, and complete the comparison table activity.

SUMMARY
A cell is the basic unit of life. Cells contain organelles that perform specific functions. Plant cells differ from animal cells in having a cell wall, chloroplasts, and a large central vacuole. The nucleus controls all cell activities while the mitochondria provides energy.

ASSIGNMENT
1. Make a large, well-labelled diagram of a plant cell in your biology notebook.
2. Write short notes on the functions of: (a) ribosome (b) cell membrane (c) nucleus.
3. Research one disease caused by cell malfunction and write half a page about it.

APPENDIX
Figure 1: Diagram of an Animal Cell showing nucleus, cell membrane, cytoplasm, mitochondria, ribosome, endoplasmic reticulum, Golgi apparatus, and lysosome.

Figure 2: Diagram of a Plant Cell showing all organelles in the animal cell plus cell wall, chloroplast, and large central vacuole.

Reference: Senior Secondary Biology by Ramalingam, Chapter 1.`;
  
  renderWeekList();
  renderLesson(1);
  document.getElementById('download-bar-wrap').style.display = 'block';
  document.getElementById('dl-summary').textContent = 'Sample lesson note – Biology SS 2, Week 1';
  document.getElementById('download-section').style.display = 'block';
  document.getElementById('generation-progress').style.display = 'none';
}

// ═══════════════════════════════════════════════════════
//  UI HELPERS
// ═══════════════════════════════════════════════════════
function showUploadProgress(show, text='', pct=0) {
  const el = document.getElementById('upload-progress');
  el.style.display = show ? 'block' : 'none';
  if(text) document.getElementById('upload-status').textContent = text;
  document.getElementById('progress-fill').style.width = pct + '%';
}

function showAlert(containerId, type, msg) {
  document.getElementById(containerId).innerHTML = `
    <div class="alert alert-${type}">${msg}</div>
  `;
}
function clearAlert(containerId) {
  document.getElementById(containerId).innerHTML = '';
}

function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
</script>

<!-- DOCX library fallback shim -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/docx/8.5.0/docx.umd.min.js"></script>
</body>
</html>