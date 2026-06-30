<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Carrefour — Tableau de bord Formations</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.9.0/dist/tabler-icons.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
:root{--blue:#002D72;--blue-mid:#185FA5;--blue-light:#E6F1FB;--red:#E3001B;--green:#1D9E75;--green-light:#E1F5EE;--amber:#EF9F27;--amber-light:#FAEEDA;--gray:#888780;--bg:#F5F6FA;--surface:#fff;--surface2:#F0F2F8;--border:rgba(0,45,114,0.12);--text:#0A1628;--text2:#4A5568;--text3:#8A96A8;--radius:10px;--radius-lg:16px;--font:'DM Sans',sans-serif;--mono:'DM Mono',monospace}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:var(--font);background:var(--bg);color:var(--text);font-size:14px;min-height:100vh}
.sidebar{position:fixed;top:0;left:0;width:220px;height:100vh;background:var(--blue);display:flex;flex-direction:column;z-index:100}
.sidebar-logo{padding:20px 20px 16px;border-bottom:1px solid rgba(255,255,255,0.1)}
.sidebar-logo-inner{display:flex;align-items:center;gap:10px}
.sidebar-logo-box{width:36px;height:36px;background:#E3001B;border-radius:8px;display:flex;align-items:center;justify-content:center;font-weight:600;font-size:16px;color:#fff;flex-shrink:0}
.sidebar-logo-text{color:#fff;font-weight:600;font-size:15px;line-height:1.2}
.sidebar-logo-sub{color:rgba(255,255,255,0.5);font-size:11px}
.sidebar-nav{padding:16px 12px;flex:1;display:flex;flex-direction:column;gap:2px;overflow-y:auto}
.nav-section{font-size:10px;font-weight:600;color:rgba(255,255,255,0.35);letter-spacing:.08em;text-transform:uppercase;padding:12px 8px 6px;margin-top:4px}
.nav-item{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;cursor:pointer;color:rgba(255,255,255,0.65);font-size:13px;transition:all .15s}
.nav-item:hover{background:rgba(255,255,255,0.08);color:#fff}
.nav-item.active{background:rgba(255,255,255,0.15);color:#fff;font-weight:500}
.nav-item i{font-size:16px;flex-shrink:0}
.sidebar-footer{padding:12px;border-top:1px solid rgba(255,255,255,0.1)}
.import-btn{width:100%;display:flex;align-items:center;gap:8px;padding:9px 12px;background:var(--red);color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:12px;font-weight:500;font-family:var(--font)}
.import-btn:hover{background:#c8001a}
.import-meta{padding:8px 4px 0;font-size:10px;color:rgba(255,255,255,0.35);line-height:1.5}
.import-meta strong{color:rgba(255,255,255,0.6);font-weight:500}
.main{margin-left:220px;min-height:100vh;display:flex;flex-direction:column}
.topbar{background:var(--surface);border-bottom:1px solid var(--border);padding:0 28px;display:flex;align-items:center;justify-content:space-between;height:56px;position:sticky;top:0;z-index:50}
.topbar-title{font-size:16px;font-weight:600;color:var(--text)}
.topbar-right{display:flex;align-items:center;gap:8px}
.topbar-date{font-size:12px;color:var(--text3);background:var(--surface2);padding:4px 10px;border-radius:20px}
.btn{display:inline-flex;align-items:center;gap:6px;padding:7px 14px;border-radius:8px;font-size:12px;font-weight:500;cursor:pointer;border:1px solid var(--border);background:var(--surface);color:var(--text);font-family:var(--font);transition:all .15s}
.btn:hover{background:var(--surface2)}
.btn-primary{background:var(--blue);color:#fff;border-color:var(--blue)}
.btn-green{background:var(--green);color:#fff;border-color:var(--green)}
.content{padding:24px 28px;flex:1}
#uploadPage{display:flex;align-items:center;justify-content:center;min-height:calc(100vh - 56px)}
.upload-card{background:var(--surface);border-radius:var(--radius-lg);border:2px dashed var(--border);padding:60px 48px;text-align:center;max-width:480px;width:100%}
.upload-card.drag{border-color:var(--blue)}
.upload-icon{width:64px;height:64px;background:var(--blue-light);border-radius:16px;display:flex;align-items:center;justify-content:center;margin:0 auto 20px;font-size:28px;color:var(--blue)}
.upload-title{font-size:20px;font-weight:600;margin-bottom:8px}
.upload-sub{font-size:13px;color:var(--text2);margin-bottom:24px;line-height:1.6}
.upload-hint{font-size:11px;color:var(--text3);margin-top:16px}
#loadingPage{display:none;align-items:center;justify-content:center;min-height:calc(100vh - 56px)}
.loading-inner{text-align:center}
.loading-spinner{width:48px;height:48px;border:3px solid var(--surface2);border-top-color:var(--blue);border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 20px}
@keyframes spin{to{transform:rotate(360deg)}}
.loading-text{font-size:14px;color:var(--text2)}
.page{display:none}.page.active{display:block}
.metrics-section{margin-bottom:20px}
.metrics-section-title{font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px;display:flex;align-items:center;gap:6px}
.metrics-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px}
.metric-card{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:14px 16px;position:relative;overflow:hidden}
.metric-card::before{content:'';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--blue)}
.metric-card.green::before{background:var(--green)}.metric-card.red::before{background:var(--red)}
.metric-card.amber::before{background:var(--amber)}.metric-card.purple::before{background:#7C3AED}
.metric-card.teal::before{background:#0891B2}.metric-card.gray::before{background:var(--gray)}
.metric-label{font-size:10px;color:var(--text3);font-weight:600;text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px}
.metric-value{font-size:24px;font-weight:700;color:var(--text);font-family:var(--mono)}
.metric-sub{font-size:10px;color:var(--text3);margin-top:3px}
.date-filter-bar{background:var(--blue-light);border:1px solid rgba(0,45,114,0.2);border-radius:var(--radius);padding:10px 16px;margin-bottom:8px;display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.date-filter-bar label{font-size:12px;font-weight:600;color:var(--blue);display:flex;align-items:center;gap:6px}
.date-input{font-size:12px;border:1px solid rgba(0,45,114,0.2);border-radius:6px;padding:5px 8px;background:var(--surface);color:var(--text);font-family:var(--font)}
.filters-bar{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:12px 16px;margin-bottom:8px;display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.filter-group-title{font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.05em;width:100%;margin-top:6px;padding-top:8px;border-top:1px solid var(--border)}
.filter-group-title:first-child{margin-top:0;padding-top:0;border-top:none}
.msel-wrap{position:relative;display:inline-block}
.msel-btn{font-size:12px;border:1px solid var(--border);border-radius:7px;padding:6px 10px;background:var(--surface);color:var(--text);cursor:pointer;display:flex;align-items:center;gap:6px;min-width:150px;justify-content:space-between;font-family:var(--font);white-space:nowrap}
.msel-btn:hover{border-color:var(--blue)}
.msel-btn.has-sel{border-color:var(--blue);background:var(--blue-light);color:var(--blue)}
.msel-dd{position:absolute;top:calc(100% + 4px);left:0;min-width:240px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);box-shadow:0 8px 24px rgba(0,0,0,0.12);z-index:300;display:none;overflow:hidden}
.msel-dd.open{display:block}
.msel-actions{display:flex;gap:4px;padding:6px 10px;border-bottom:1px solid var(--border);background:var(--surface2)}
.msel-actions button{flex:1;font-size:11px;padding:3px 6px;border-radius:5px;border:1px solid var(--border);background:var(--surface);color:var(--text);cursor:pointer;font-family:var(--font)}
.msel-actions button:hover{background:var(--blue);color:#fff;border-color:var(--blue)}
.msel-search-inp{width:calc(100% - 20px);margin:6px 10px;border:1px solid var(--border);border-radius:6px;padding:5px 8px;font-size:12px;font-family:var(--font);color:var(--text);background:var(--surface2);display:block}
.msel-search-inp:focus{outline:none;border-color:var(--blue)}
.msel-list{max-height:240px;overflow-y:auto}
.msel-opt{display:flex;align-items:center;gap:8px;padding:7px 12px;cursor:pointer;font-size:12px}
.msel-opt:hover{background:var(--surface2)}
.msel-opt input[type=checkbox]{accent-color:var(--blue);width:13px;height:13px;cursor:pointer;flex-shrink:0}
.chips-row{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:10px;min-height:4px}
.chip{display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:500}
.chip-blue{background:var(--blue);color:#fff}.chip-green{background:var(--green);color:#fff}
.chip-amber{background:var(--amber);color:#fff}.chip-red{background:var(--red);color:#fff}
.chip-gray{background:var(--gray);color:#fff}.chip-teal{background:#0891B2;color:#fff}.chip-purple{background:#7C3AED;color:#fff}
.chip-x{cursor:pointer;opacity:.7;font-size:13px}.chip-x:hover{opacity:1}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}
.card{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:20px}
.card-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.card-title{font-size:13px;font-weight:600;color:var(--text)}
.badge{display:inline-flex;align-items:center;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600}
.badge-green{background:var(--green-light);color:#0F6E56}.badge-amber{background:var(--amber-light);color:#854F0B}
.badge-red{background:#FCEBEB;color:#A32D2D}.badge-blue{background:var(--blue-light);color:var(--blue-mid)}
.badge-gray{background:var(--surface2);color:var(--text2)}.badge-purple{background:#EDE9FE;color:#5B21B6}
.badge-teal{background:#CFFAFE;color:#0E7490}
.tabs{display:flex;border-bottom:1px solid var(--border);margin-bottom:16px;flex-wrap:wrap}
.tab{font-size:12px;padding:8px 12px;cursor:pointer;color:var(--text3);border-bottom:2px solid transparent;margin-bottom:-1px;font-weight:500;white-space:nowrap;display:flex;align-items:center;gap:5px}
.tab.active{color:var(--blue);border-bottom-color:var(--blue)}
.tab-pane{display:none}.tab-pane.active{display:block}
.tbl-wrap{overflow-x:auto;border-radius:8px;border:1px solid var(--border)}
table{width:100%;border-collapse:collapse;font-size:12px}
th{text-align:left;padding:9px 12px;background:var(--surface2);font-weight:600;color:var(--text2);font-size:11px;text-transform:uppercase;letter-spacing:.03em;border-bottom:1px solid var(--border)}
td{padding:8px 12px;border-bottom:1px solid var(--border);color:var(--text)}
tr:last-child td{border-bottom:none}
tr:hover td{background:#FAFBFF}
tr.selected td{background:#EFF6FF}
.td-name{font-weight:500;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cb-col{width:32px;padding:8px!important}
input[type=checkbox].row-cb{accent-color:var(--blue);width:14px;height:14px;cursor:pointer}
.prog-row{display:flex;align-items:center;gap:8px;margin-bottom:7px}
.prog-label{font-size:12px;color:var(--text);min-width:180px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex-shrink:0}
.prog-bg{flex:1;height:6px;background:var(--surface2);border-radius:3px;overflow:hidden}
.prog-fill{height:100%;border-radius:3px}
.prog-count{font-size:11px;color:var(--text3);min-width:44px;text-align:right;font-family:var(--mono)}
.search-wrap{position:relative;margin-bottom:16px}
.search-ico{position:absolute;left:11px;top:50%;transform:translateY(-50%);color:var(--text3);font-size:16px}
.search-input{width:100%;padding:9px 12px 9px 36px;border:1px solid var(--border);border-radius:8px;font-size:13px;font-family:var(--font);color:var(--text);background:var(--surface)}
.search-input:focus{outline:none;border-color:var(--blue)}
.suggestions{position:absolute;top:calc(100% + 4px);left:0;right:0;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);z-index:100;max-height:280px;overflow-y:auto;display:none;box-shadow:0 8px 24px rgba(0,0,0,0.1)}
.sug-item{padding:10px 14px;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:space-between}
.sug-item:hover,.sug-item.active{background:var(--surface2)}
.sug-name{font-weight:500}.sug-meta{font-size:11px;color:var(--text3)}
.search-clear{position:absolute;right:10px;top:50%;transform:translateY(-50%);color:var(--text3);font-size:16px;cursor:pointer;display:none}
.fiche-header{background:linear-gradient(135deg,var(--blue) 0%,#185FA5 100%);color:#fff;border-radius:var(--radius-lg);padding:20px 24px;margin-bottom:16px}
.fiche-name{font-size:18px;font-weight:600;margin-bottom:6px}
.fiche-tags{display:flex;flex-wrap:wrap;gap:6px}
.fiche-tag{background:rgba(255,255,255,0.2);color:#fff;padding:2px 10px;border-radius:20px;font-size:11px}
.fiche-date{font-size:11px;opacity:.6;margin-top:6px}
.fiche-section{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:16px;margin-bottom:16px}
.fiche-section-title{font-size:12px;font-weight:700;color:var(--text2);text-transform:uppercase;letter-spacing:.04em;margin-bottom:12px;display:flex;align-items:center;gap:6px}
.fiche-metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:8px}
.fiche-m{background:var(--surface2);border-radius:8px;padding:10px 12px;border-left:3px solid var(--blue)}
.fiche-m.green{border-left-color:var(--green)}.fiche-m.red{border-left-color:var(--red)}
.fiche-m.amber{border-left-color:var(--amber)}.fiche-m.purple{border-left-color:#7C3AED}
.fiche-m.teal{border-left-color:#0891B2}.fiche-m.gray{border-left-color:var(--gray)}
.fiche-m-label{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px}
.fiche-m-value{font-size:20px;font-weight:700;font-family:var(--mono)}
.fiche-m-sub{font-size:10px;margin-top:2px;color:var(--text3)}
.chart-box{position:relative;width:100%}
.empty{text-align:center;padding:40px;color:var(--text3)}
.empty i{font-size:40px;opacity:.3;display:block;margin-bottom:12px}
.info-box{background:var(--blue-light);border-radius:8px;padding:8px 12px;font-size:12px;color:var(--blue-mid);margin-bottom:12px;display:flex;align-items:flex-start;gap:6px;line-height:1.6}
.warn-box{background:var(--amber-light);border-radius:8px;padding:8px 12px;font-size:12px;color:#854F0B;margin-bottom:10px;display:flex;align-items:center;gap:6px}
.sel-actions{display:flex;gap:8px;margin-bottom:8px;align-items:center;font-size:12px;color:var(--text3);flex-wrap:wrap}
@media(max-width:900px){
.sidebar{width:64px}
.sidebar-logo-text,.nav-item span,.nav-section,.sidebar-footer{display:none}
.sidebar-logo-box{margin:0 auto}.sidebar-logo-inner{justify-content:center}
.nav-item{justify-content:center;padding:10px}
.main{margin-left:64px}
.two-col{grid-template-columns:1fr}}
</style>
</head>
<body>
<nav class="sidebar">
  <div class="sidebar-logo"><div class="sidebar-logo-inner"><div class="sidebar-logo-box">C</div><div><div class="sidebar-logo-text">Carrefour</div><div class="sidebar-logo-sub">Formations</div></div></div></div>
  <div class="sidebar-nav">
    <div class="nav-section">Vues</div>
    <div class="nav-item active" onclick="showPage('dashboard')" id="nav-dashboard"><i class="ti ti-layout-dashboard"></i><span>Tableau de bord</span></div>
    <div class="nav-item" onclick="showPage('magasin')" id="nav-magasin"><i class="ti ti-building-store"></i><span>Fiche magasin</span></div>
    <div class="nav-section">Analyse</div>
    <div class="nav-item" onclick="showPage('formations')" id="nav-formations"><i class="ti ti-school"></i><span>Formations</span></div>
    <div class="nav-item" onclick="showPage('stagiaires')" id="nav-stagiaires"><i class="ti ti-users"></i><span>Stagiaires</span></div>
  </div>
  <div class="sidebar-footer">
    <button class="import-btn" onclick="document.getElementById('fileInput').click()"><i class="ti ti-refresh" style="font-size:15px"></i><span>Mise à jour mensuelle</span></button>
    <div class="import-meta" id="importMeta">Aucun fichier chargé</div>
    <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none">
  </div>
</nav>
<div class="main">
  <div class="topbar">
    <div class="topbar-title" id="pageTitle">Tableau de bord</div>
    <div class="topbar-right">
      <span class="topbar-date" id="topbarDate">—</span>
      <button class="btn" onclick="exportCSV()"><i class="ti ti-file-text"></i> CSV</button>
      <button class="btn btn-green" onclick="exportXLSX()"><i class="ti ti-file-spreadsheet"></i> Excel</button>
    </div>
  </div>
  <div id="uploadPage" style="display:flex">
    <div class="upload-card" id="dropZone">
      <div class="upload-icon"><i class="ti ti-file-spreadsheet"></i></div>
      <div class="upload-title">Importer les données formations</div>
      <div class="upload-sub">Glisse ton fichier Excel ici, ou clique pour le sélectionner.<br>Feuilles requises : <strong>Data</strong> et <strong>Liste des DT</strong>.</div>
      <button class="btn btn-primary" style="font-size:13px;padding:10px 20px" onclick="document.getElementById('fileInput').click()"><i class="ti ti-upload"></i> Choisir le fichier</button>
      <div class="upload-hint">Formats : .xlsx · .xls</div>
    </div>
  </div>
  <div id="loadingPage" style="display:none">
    <div class="loading-inner"><div class="loading-spinner"></div><div class="loading-text" id="loadingMsg">Chargement...</div></div>
  </div>
  <div class="content" id="pagesWrap" style="display:none">
    <div class="date-filter-bar">
      <i class="ti ti-calendar-event" style="color:var(--blue);font-size:16px"></i>
      <label>Du <input type="date" id="dateFrom" class="date-input" onchange="applyFilters()"></label>
      <label>Au <input type="date" id="dateTo" class="date-input" onchange="applyFilters()"></label>
      <button class="btn" style="font-size:11px;padding:5px 10px" onclick="resetDates()"><i class="ti ti-refresh"></i> Toutes périodes</button>
      <span id="dateInfo" style="font-size:11px;color:var(--blue-mid)"></span>
    </div>
    <div class="filters-bar">
      <div class="filter-group-title" style="margin-top:0;padding-top:0;border-top:none">🌍 Géographie &amp; Client</div>
      <div class="msel-wrap"><div class="msel-btn" id="b-dt" onclick="toggleMsel('d-dt')"><span id="l-dt">Directions (DT)</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-dt"></div></div>
      <div class="msel-wrap"><div class="msel-btn" id="b-region" onclick="toggleMsel('d-region')"><span id="l-region">Régions</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-region"></div></div>
      <div class="msel-wrap"><div class="msel-btn" id="b-categorie" onclick="toggleMsel('d-categorie')"><span id="l-categorie">Catégories client</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-categorie"></div></div>
      <div class="msel-wrap"><div class="msel-btn" id="b-client" onclick="toggleMsel('d-client')" style="min-width:170px"><span id="l-client">Clients / Magasins</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-client"></div></div>
      <div class="filter-group-title">📋 Formation</div>
      <div class="msel-wrap"><div class="msel-btn" id="b-domaine" onclick="toggleMsel('d-domaine')"><span id="l-domaine">Domaines</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-domaine"></div></div>
      <div class="msel-wrap"><div class="msel-btn" id="b-sdomaine" onclick="toggleMsel('d-sdomaine')"><span id="l-sdomaine">Sous-domaines</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-sdomaine"></div></div>
      <div class="msel-wrap"><div class="msel-btn" id="b-formation" onclick="toggleMsel('d-formation')" style="min-width:170px"><span id="l-formation">Thèmes</span><i class="ti ti-chevron-down" style="font-size:11px;opacity:.5"></i></div><div class="msel-dd" id="d-formation"></div></div>
      <button class="btn" onclick="resetFilters()" style="margin-left:auto"><i class="ti ti-refresh"></i> Reset tout</button>
    </div>
    <div class="chips-row" id="allChips"></div>

    <div class="page active" id="page-dashboard">
      <div class="info-box"><i class="ti ti-info-circle" style="flex-shrink:0"></i><div>
        <strong>Règles de qualification (col AB + col C + Date de fin) :</strong><br>
        • Statut session = Annulée → <strong>Annulé</strong> (priorité absolue)<br>
        • Qualification = Présent / Présent partiel → <strong>Réalisé</strong><br>
        • Qualification = Absent → <strong>Absent</strong><br>
        • Qualification = Annulé → <strong>Annulé</strong><br>
        • Qualification vide + (Clôturée/En cours/Validée) + date fin ≤ aujourd'hui → <strong>Réalisé</strong><br>
        • Qualification vide + date fin &gt; aujourd'hui → <strong>Inscrit</strong>
      </div></div>
      <div class="metrics-section">
        <div class="metrics-section-title"><i class="ti ti-calendar" style="font-size:14px"></i> Sessions (col. A &amp; C)</div>
        <div class="metrics-grid" id="metricsSessionsRow"></div>
      </div>
      <div class="metrics-section">
        <div class="metrics-section-title"><i class="ti ti-users" style="font-size:14px"></i> Stagiaires (col. V, AB &amp; date fin)</div>
        <div class="metrics-grid" id="metricsStagiairesRow"></div>
      </div>
      <div class="two-col">
        <div>
          <div class="card" style="margin-bottom:16px"><div class="card-header"><div class="card-title">Évolution mensuelle</div></div><div class="chart-box" style="height:230px"><canvas id="chartMensuel"></canvas></div></div>
          <div class="card"><div class="card-header"><div class="card-title">Répartition par domaine (Réalisés)</div></div><div id="dashDomaines"></div></div>
        </div>
        <div>
          <div class="card" style="margin-bottom:16px"><div class="card-header"><div class="card-title">Statut des sessions (col C)</div></div><div class="chart-box" style="height:160px"><canvas id="chartStatut"></canvas></div><div id="statutLegend" style="margin-top:8px"></div></div>
          <div class="card"><div class="card-header"><div class="card-title">Qualification stagiaires</div></div><div class="chart-box" style="height:140px"><canvas id="chartQualif"></canvas></div><div id="qualifLegend" style="margin-top:8px"></div></div>
        </div>
      </div>
    </div>

    <div class="page" id="page-magasin">
      <div class="search-wrap"><i class="ti ti-building-store search-ico"></i><input type="text" class="search-input" id="searchMag" placeholder="Rechercher un magasin / client..." autocomplete="off"><i class="ti ti-x search-clear" id="searchClear" onclick="clearSearch()"></i><div class="suggestions" id="suggestions"></div></div>
      <div id="ficheContent" style="display:none">
        <div class="fiche-header"><div class="fiche-name" id="ficheName"></div><div class="fiche-tags" id="ficheTags"></div><div class="fiche-date" id="ficheDate"></div></div>
        <div class="fiche-section">
          <div class="fiche-section-title"><i class="ti ti-calendar"></i> Sessions</div>
          <div class="fiche-metrics" id="ficheMetricsSess"></div>
        </div>
        <div class="fiche-section">
          <div class="fiche-section-title"><i class="ti ti-users"></i> Stagiaires</div>
          <div class="fiche-metrics" id="ficheMetricsStag"></div>
        </div>
        <div class="tabs" id="ficheTabs">
          <div class="tab active" onclick="switchFicheTab('sessions')"><i class="ti ti-calendar"></i> Sessions <span id="cnt-sess" class="badge badge-green">0</span></div>
          <div class="tab" onclick="switchFicheTab('inscrits')"><i class="ti ti-calendar-event"></i> Inscrits <span id="cnt-inscrit" class="badge badge-teal">0</span></div>
          <div class="tab" onclick="switchFicheTab('stagiaires')"><i class="ti ti-users"></i> Stagiaires <span id="cnt-stag" class="badge badge-blue">0</span></div>
          <div class="tab" onclick="switchFicheTab('annules')"><i class="ti ti-x"></i> Annulés <span id="cnt-ann" class="badge badge-red">0</span></div>
        </div>
        <div id="fiche-tab-sessions" class="tab-pane active">
          <div class="tbl-wrap"><table>
            <thead><tr><th>Formation</th><th>Domaine</th><th>Date début</th><th>Date fin</th><th>Statut</th></tr></thead>
            <tbody id="sessTableBody"></tbody>
          </table></div>
        </div>
        <div id="fiche-tab-inscrits" class="tab-pane">
          <div class="tbl-wrap"><table>
            <thead><tr><th>Stagiaire</th><th>Formation</th><th>Date Début</th><th>Date Fin</th></tr></thead>
            <tbody id="insTableBody"></tbody>
          </table></div>
        </div>
        <div id="fiche-tab-stagiaires" class="tab-pane">
          <div class="tbl-wrap"><table>
            <thead><tr><th>Matricule</th><th>Nom Prénom</th><th>Formation</th><th>Statut</th></tr></thead>
            <tbody id="stagTableBody"></tbody>
          </table></div>
        </div>
        <div id="fiche-tab-annules" class="tab-pane">
          <div class="tbl-wrap"><table>
            <thead><tr><th>Nom Prénom</th><th>Formation</th><th>Type Annulation</th></tr></thead>
            <tbody id="annTableBody"></tbody>
          </table></div>
        </div>
      </div>
      <div class="empty" id="ficheEmpty"><i class="ti ti-building-store"></i><p>Recherche un magasin pour afficher sa fiche</p></div>
    </div>

    <div class="page" id="page-formations">
      <div class="card" style="margin-bottom:16px"><div class="card-header"><div class="card-title">Top formations (Réalisés)</div></div><div id="topFormationsProgress"></div></div>
      <div class="card"><div class="card-header"><div class="card-title">Toutes les formations</div></div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Formation</th><th>Domaine</th><th>Réalisés</th><th>Inscrits</th></tr></thead>
          <tbody id="tableFormationsBody"></tbody>
        </table></div>
      </div>
    </div>

    <div class="page" id="page-stagiaires">
      <div class="card">
        <div class="search-wrap"><i class="ti ti-users search-ico"></i><input type="text" class="search-input" id="searchStagGlobal" placeholder="Filtrer par nom ou matricule..."></div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Matricule</th><th>Nom</th><th>Prénom</th><th>Magasin</th><th>Formation</th><th>Statut</th></tr></thead>
          <tbody id="tableStagiairesBody"></tbody>
        </table></div>
      </div>
    </div>

  </div>
</div>

<script>
// VARIABLES GLOBALES ET STRUCTURES DE DONNÉES
let rawData = [];
let filteredData = [];
let selectedFilters = { dt:[], region:[], categorie:[], client:[], domaine:[], sdomaine:[], formation:[] };
let charts = {};

document.getElementById('fileInput').addEventListener('change', handleFile, false);
setupDragAndDrop();

function setupDragAndDrop() {
  let dz = document.getElementById('dropZone');
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
  dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag'); if(e.dataTransfer.files.length) { document.getElementById('fileInput').files = e.dataTransfer.files; handleFile(); } });
}

function handleFile() {
  let file = document.getElementById('fileInput').files[0];
  if(!file) return;
  document.getElementById('uploadPage').style.display = 'none';
  document.getElementById('loadingPage').style.display = 'flex';
  
  let reader = new FileReader();
  reader.onload = function(e) {
    let data = new Uint8Array(e.target.result);
    let workbook = XLSX.read(data, {type: 'array', cellDates: true});
    let sheetName = workbook.SheetNames.includes("Data") ? "Data" : workbook.SheetNames[0];
    let sheet = workbook.Sheets[sheetName];
    
    // Extraction brute (sans headers restrictifs pour plus de souplesse)
    let json = XLSX.utils.sheet_to_json(sheet, {header: 1});
    processData(json);
  };
  reader.readAsArrayBuffer(file);
}

function processData(rows) {
  if(rows.length < 2) { alert("Fichier vide ou invalide"); return; }
  rawData = [];
  let now = new Date();

  // Mappage simplifié basé sur les lettres de colonnes spécifiées dans l'énoncé original
  // Col A=0 (ID Session), C=2 (Statut Session), V=21 (Nom/Magasin), AB=27 (Qualification), etc.
  for(let i=1; i<rows.length; i++) {
    let r = rows[i];
    if(!r[0]) continue; // Saute les lignes vides

    let dateFin = r[4] instanceof Date ? r[4] : new Date(r[4] || now);
    let dateDeb = r[3] instanceof Date ? r[3] : new Date(r[3] || now);
    let statutSess = String(r[2] || '').trim();
    let qualif = String(r[27] || '').trim();
    
    // Application stricte de vos règles de qualification
    let statutStagiaire = "Inscrit";
    if (statutSess.toLowerCase() === "annulée" || qualif.toLowerCase() === "annulé") {
      statutStagiaire = "Annulé";
    } else if (qualif.toLowerCase() === "présent" || qualif.toLowerCase() === "présent partiel") {
      statutStagiaire = "Réalisé";
    } else if (qualif.toLowerCase() === "absent") {
      statutStagiaire = "Absent";
    } else if (qualif === "") {
      if (["clôturée", "en cours", "validée"].includes(statutSess.toLowerCase()) && dateFin <= now) {
        statutStagiaire = "Réalisé";
      } else if (dateFin > now) {
        statutStagiaire = "Inscrit";
      }
    }

    rawData.push({
      idSession: r[0],
      theme: r[1] || 'Non spécifié',
      statutSession: statutSess || 'Inconnu',
      dateDebut: dateDeb,
      dateFin: dateFin,
      domaine: r[5] || 'Général',
      sousDomaine: r[6] || 'Général',
      magasin: r[21] || 'Siège / Autre',
      matricule: r[22] || 'Inconnu',
      nom: r[23] || 'Anonyme',
      prenom: r[24] || '',
      dt: r[18] || 'Hors DT',
      region: r[19] || 'Autre Région',
      categorie: r[20] || 'Standard',
      statutStagiaire: statutStagiaire
    });
  }

  document.getElementById('importMeta').innerHTML = `<strong>${rawData.length}</strong> lignes chargées`;
  document.getElementById('topbarDate').innerText = now.toLocaleDateString('fr-FR');
  document.getElementById('loadingPage').style.display = 'none';
  document.getElementById('pagesWrap').style.display = 'block';

  buildFiltersDropdowns();
  applyFilters();
}

function buildFiltersDropdowns() {
  let unique = (key) => [...new Set(rawData.map(d => d[key]))].filter(Boolean).sort();
  populateDropdown('d-dt', unique('dt'), 'dt');
  populateDropdown('d-region', unique('region'), 'region');
  populateDropdown('d-categorie', unique('categorie'), 'categorie');
  populateDropdown('d-client', unique('magasin'), 'client');
  populateDropdown('d-domaine', unique('domaine'), 'domaine');
  populateDropdown('d-sdomaine', unique('sousDomaine'), 'sdomaine');
  populateDropdown('d-formation', unique('theme'), 'formation');
}

function populateDropdown(id, list, key) {
  let div = document.getElementById(id);
  let html = `<div class="msel-actions"><button onclick="checkAllMsel('${key}', true)">Tous</button><button onclick="checkAllMsel('${key}', false)">Aucun</button></div>`;
  html += `<input type="text" class="msel-search-inp" placeholder="Rechercher..." oninput="filterMselList('${id}', this.value)">`;
  html += `<div class="msel-list">`;
  list.forEach(val => {
    html += `<label class="msel-opt"><input type="checkbox" value="${val}" onchange="handleMselChange('${key}')"> ${val}</label>`;
  });
  html += `</div>`;
  div.innerHTML = html;
}

function toggleMsel(id) {
  let el = document.getElementById(id);
  let open = el.style.display === 'block';
  document.querySelectorAll('.msel-dd').forEach(d => d.style.display = 'none');
  el.style.display = open ? 'none' : 'block';
}

function handleMselChange(key) {
  let checked = [...document.querySelectorAll(`#d-${key} input[type=checkbox]:checked`)].map(cb => cb.value);
  selectedFilters[key] = checked;
  let btn = document.getElementById(`b-${key}`);
  if(checked.length) { btn.classList.add('has-sel'); } else { btn.classList.remove('has-sel'); }
  applyFilters();
}

function applyFilters() {
  filteredData = rawData.filter(d => {
    if(selectedFilters.dt.length && !selectedFilters.dt.includes(d.dt)) return false;
    if(selectedFilters.region.length && !selectedFilters.region.includes(d.region)) return false;
    if(selectedFilters.categorie.length && !selectedFilters.categorie.includes(d.categorie)) return false;
    if(selectedFilters.client.length && !selectedFilters.client.includes(d.magasin)) return false;
    if(selectedFilters.domaine.length && !selectedFilters.domaine.includes(d.domaine)) return false;
    if(selectedFilters.sdomaine.length && !selectedFilters.sdomaine.includes(d.sousDomaine)) return false;
    if(selectedFilters.formation.length && !selectedFilters.formation.includes(d.theme)) return false;
    return true;
  });
  updateDashboard();
}

function updateDashboard() {
  // Calculs KPIs
  let totalSessions = [...new Set(filteredData.map(d => d.idSession))].length;
  let totalStagiaires = filteredData.length;
  let realises = filteredData.filter(d => d.statutStagiaire === 'Réalisé').length;
  let inscrits = filteredData.filter(d => d.statutStagiaire === 'Inscrit').length;
  let absents = filteredData.filter(d => d.statutStagiaire === 'Absent').length;
  let annules = filteredData.filter(d => d.statutStagiaire === 'Annulé').length;

  // Injection HTML des blocs KPIs
  document.getElementById('metricsSessionsRow').innerHTML = `
    <div class="metric-card"><div class="metric-label">Total Sessions</div><div class="metric-value">${totalSessions}</div></div>
  `;
  document.getElementById('metricsStagiairesRow').innerHTML = `
    <div class="metric-card green"><div class="metric-label">Réalisés</div><div class="metric-value">${realises}</div></div>
    <div class="metric-card teal"><div class="metric-label">Inscrits</div><div class="metric-value">${inscrits}</div></div>
    <div class="metric-card amber"><div class="metric-label">Absents</div><div class="metric-value">${absents}</div></div>
    <div class="metric-card red"><div class="metric-label">Annulés</div><div class="metric-value">${annules}</div></div>
  `;

  renderCharts(realises, inscrits, absents, annules);
  updateFormationsPage();
  updateStagiairesPage();
  buildMagasinAutocomplete();
}

function renderCharts(r, i, a, an) {
  // Graphique Qualif Stagiaires
  if(charts.qualif) charts.qualif.destroy();
  let ctx = document.getElementById('chartQualif').getContext('2d');
  charts.qualif = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Réalisé', 'Inscrit', 'Absent', 'Annulé'],
      datasets: [{ data: [r, i, a, an], backgroundColor: ['#1D9E75', '#0891B2', '#EF9F27', '#E3001B'] }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
  });
}

function updateFormationsPage() {
  let counts = {};
  filteredData.forEach(d => { counts[d.theme] = (counts[d.theme] || 0) + (d.statutStagiaire === 'Réalisé' ? 1 : 0); });
  let sorted = Object.entries(counts).sort((a,b) => b[1] - a[1]);

  let html = '';
  sorted.slice(0, 5).forEach(([theme, val]) => {
    html += `
      <div class="prog-row">
        <div class="prog-label">${theme}</div>
        <div class="prog-bg"><div class="prog-fill" style="width:${Math.min(val*10, 100)}%; background:var(--blue)"></div></div>
        <div class="prog-count">${val}</div>
      </div>`;
  });
  document.getElementById('topFormationsProgress').innerHTML = html || 'Aucune donnée';
}

function buildMagasinAutocomplete() {
  let mags = [...new Set(filteredData.map(d => d.magasin))];
  let inp = document.getElementById('searchMag');
  let sug = document.getElementById('suggestions');
  
  inp.oninput = function() {
    let val = this.value.toLowerCase();
    sug.innerHTML = '';
    if(!val) { sug.style.display = 'none'; return; }
    let matches = mags.filter(m => m.toLowerCase().includes(val));
    if(matches.length) {
      sug.style.display = 'block';
      matches.forEach(m => {
        let d = document.createElement('div');
        d.className = 'sug-item';
        d.innerText = m;
        d.onclick = () => { inp.value = m; sug.style.display = 'none'; showMagasinFiche(m); };
        sug.appendChild(d);
      });
    } else { sug.style.display = 'none'; }
  };
}

function showMagasinFiche(magName) {
  document.getElementById('ficheEmpty').style.display = 'none';
  let content = document.getElementById('ficheContent');
  content.style.display = 'block';
  document.getElementById('ficheName').innerText = magName;

  let magData = filteredData.filter(d => d.magasin === magName);
  
  // Remplissage des sous-tableaux de la fiche magasin
  let sBody = document.getElementById('sessTableBody'); sBody.innerHTML = '';
  let iBody = document.getElementById('insTableBody'); iBody.innerHTML = '';
  
  let uniqueSessions = [...new Set(magData.map(d => d.idSession))];
  document.getElementById('cnt-sess').innerText = uniqueSessions.length;
  document.getElementById('cnt-inscrit').innerText = magData.filter(d => d.statutStagiaire === 'Inscrit').length;
  document.getElementById('cnt-stag').innerText = magData.filter(d => d.statutStagiaire === 'Réalisé').length;
  document.getElementById('cnt-ann').innerText = magData.filter(d => d.statutStagiaire === 'Annulé').length;

  magData.forEach(d => {
    if(d.statutStagiaire === 'Inscrit') {
      iBody.innerHTML += `<tr><td>${d.nom} ${d.prenom}</td><td>${d.theme}</td><td>${d.dateDebut.toLocaleDateString('fr-FR')}</td><td>${d.dateFin.toLocaleDateString('fr-FR')}</td></tr>`;
    }
    sBody.innerHTML += `<tr><td>${d.theme}</td><td>${d.domaine}</td><td>${d.dateDebut.toLocaleDateString('fr-FR')}</td><td>${d.dateFin.toLocaleDateString('fr-FR')}</td><td>${d.statutSession}</td></tr>`;
  });
}

function updateStagiairesPage() {
  let body = document.getElementById('tableStagiairesBody');
  body.innerHTML = '';
  filteredData.slice(0, 100).forEach(d => { // Limité à 100 pour la performance fluide
    body.innerHTML += `<tr><td>${d.matricule}</td><td>${d.nom}</td><td>${d.prenom}</td><td>${d.magasin}</td><td>${d.theme}</td><td>${d.statutStagiaire}</td></tr>`;
  });
}

function showPage(pageId) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + pageId).classList.add('active');
  document.getElementById('nav-' + pageId).classList.add('active');
  document.getElementById('pageTitle').innerText = document.getElementById('nav-' + pageId).querySelector('span').innerText;
}

function resetFilters() {
  selectedFilters = { dt:[], region:[], categorie:[], client:[], domaine:[], sdomaine:[], formation:[] };
  document.querySelectorAll('.msel-btn').forEach(b => b.classList.remove('has-sel'));
  document.querySelectorAll('.msel-dd input[type=checkbox]').forEach(cb => cb.checked = false);
  applyFilters();
}
</script>
</body>
</html>
