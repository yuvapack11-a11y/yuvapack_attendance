<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>YUVAPACK — Employee Attendance Management System</title>

<!-- Professional typography -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">

<!-- External libraries (allowed by spec): charts, Excel export, PDF export -->
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>

<style>
/* ============================================================
   YUVAPACK — Design System
   ============================================================ */
:root{
  --primary:#2196F3;
  --primary-dark:#1976D2;
  --primary-darker:#0D47A1;
  --primary-light:#BBDEFB;
  --secondary:#E3F2FD;
  --accent:#499A13;
  --accent-light:#66BB6A;
  --accent-dark:#388E3C;
  --white:#FFFFFF;
  --bg:#F5F7FA;
  --bg-alt:#EEF3F8;
  --grey-light:#ECEFF1;
  --grey:#B0BEC5;
  --grey-dark:#78909C;
  --text:#263238;
  --text-light:#607D8B;
  --danger:#E53935;
  --danger-light:#FFEBEE;
  --warning:#FB8C00;
  --warning-light:#FFF3E0;
  --info:#00ACC1;
  --shadow:0 2px 8px rgba(33,150,243,.08),0 1px 3px rgba(0,0,0,.06);
  --shadow-md:0 4px 16px rgba(33,150,243,.10),0 2px 6px rgba(0,0,0,.06);
  --shadow-lg:0 12px 32px rgba(33,150,243,.16),0 4px 12px rgba(0,0,0,.08);
  --radius:16px;
  --radius-sm:10px;
  --radius-xs:6px;
  --sidebar-w:260px;
  --transition:.22s cubic-bezier(.4,0,.2,1);
}

*{margin:0;padding:0;box-sizing:border-box;}
html,body{height:100%;}
body{
  font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
  background:var(--bg);
  color:var(--text);
  font-size:14px;
  line-height:1.5;
  -webkit-font-smoothing:antialiased;
}
button{font-family:inherit;cursor:pointer;border:none;background:none;}
input,select,textarea{font-family:inherit;font-size:14px;}
a{color:var(--primary);text-decoration:none;}
::-webkit-scrollbar{width:10px;height:10px;}
::-webkit-scrollbar-thumb{background:var(--grey);border-radius:8px;border:2px solid var(--bg);}
::-webkit-scrollbar-thumb:hover{background:var(--grey-dark);}
::-webkit-scrollbar-track{background:transparent;}

/* ============================================================
   LOGIN SCREEN
   ============================================================ */
#loginScreen {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;

  /* Background Image + Gradient Overlay */
  background:
    linear-gradient(
      135deg,
      rgba(245, 247, 250, 0.85),
      rgba(227, 242, 253, 0.75)
    ),
    url("https://png.pngtree.com/thumb_back/fw800/background/20251009/pngtree-modern-office-space-with-glass-walls-and-city-view-image_19774423.webp");

  background-size: cover;
  background-position: center;
  background-repeat: no-repeat;
  background-attachment: fixed;

  padding: 24px;
  position: relative;
  overflow: hidden;
}
#loginScreen::before{
  content:"";position:absolute;inset:0;
  background-image:
    radial-gradient(circle at 20% 80%, rgba(33,150,243,.06) 0%, transparent 40%),
    radial-gradient(circle at 80% 20%, rgba(73,154,19,.05) 0%, transparent 40%);
  pointer-events:none;
}
.login-card{
  position:relative;z-index:2;
  background:var(--white);
  border-radius:24px;
  box-shadow:var(--shadow-lg);
  width:100%;max-width:440px;
  padding:40px 36px;
  animation:rise .5s cubic-bezier(.4,0,.2,1);
}
@keyframes rise{from{opacity:0;transform:translateY(20px);}to{opacity:1;transform:translateY(0);}}
.login-logo{
  display:flex;flex-direction:column;align-items:center;gap:10px;margin-bottom:6px;
}
.logo-mark{
  width:64px;height:64px;border-radius:18px;
  background:linear-gradient(135deg,var(--primary) 0%,var(--primary-dark) 100%);
  display:flex;align-items:center;justify-content:center;
  color:#fff;font-weight:800;font-size:26px;letter-spacing:.5px;
  box-shadow:0 8px 20px rgba(33,150,243,.35);
}
.login-logo h1{font-size:24px;font-weight:800;color:var(--text);letter-spacing:.5px;}
.login-logo p{font-size:13px;color:var(--text-light);text-align:center;}
.login-tabs{
  display:flex;background:var(--bg);border-radius:12px;padding:4px;margin:22px 0;
}
.login-tab{
  flex:1;padding:10px;border-radius:9px;font-weight:600;font-size:13px;
  color:var(--text-light);transition:var(--transition);
}
.login-tab.active{
  background:var(--white);color:var(--primary);
  box-shadow:0 2px 8px rgba(0,0,0,.08);
}
.form-group{margin-bottom:18px;}
.form-group label{
  display:block;font-size:12px;font-weight:600;color:var(--text-light);
  margin-bottom:7px;letter-spacing:.3px;text-transform:uppercase;
}
.input-wrap{position:relative;}
.input-wrap .icon{
  position:absolute;left:14px;top:50%;transform:translateY(-50%);
  color:var(--grey-dark);width:18px;height:18px;
}
.form-input{
  width:100%;padding:13px 14px 13px 42px;
  border:1.5px solid var(--grey-light);border-radius:var(--radius-sm);
  background:var(--white);transition:var(--transition);color:var(--text);
}
.form-input:focus{
  outline:none;border-color:var(--primary);
  box-shadow:0 0 0 4px rgba(33,150,243,.12);
}
.form-input::placeholder{color:var(--grey);}
.btn{
  display:inline-flex;align-items:center;justify-content:center;gap:8px;
  padding:12px 20px;border-radius:var(--radius-sm);font-weight:600;font-size:14px;
  transition:var(--transition);white-space:nowrap;
}
.btn-primary{
  background:var(--primary);color:#fff;width:100%;
  box-shadow:0 4px 12px rgba(33,150,243,.3);
}
.btn-primary:hover{background:var(--primary-dark);transform:translateY(-1px);box-shadow:0 6px 18px rgba(33,150,243,.4);}
.btn-primary:active{transform:translateY(0);}
.btn-accent{background:var(--accent);color:#fff;box-shadow:0 4px 12px rgba(73,154,19,.3);}
.btn-accent:hover{background:var(--accent-dark);transform:translateY(-1px);}
.btn-ghost{background:var(--bg);color:var(--text);}
.btn-ghost:hover{background:var(--grey-light);}
.btn-danger{background:var(--danger);color:#fff;}
.btn-danger:hover{background:#C62828;}
.btn-outline{border:1.5px solid var(--primary);color:var(--primary);background:transparent;}
.btn-outline:hover{background:var(--secondary);}
.btn-sm{padding:8px 14px;font-size:12.5px;border-radius:8px;}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none!important;}
.demo-creds{
  margin-top:22px;padding:14px 16px;background:var(--secondary);border-radius:var(--radius-sm);
  border-left:3px solid var(--primary);
}
.demo-creds h4{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--primary-dark);margin-bottom:8px;}
.demo-creds p{font-size:12px;color:var(--text);line-height:1.7;}
.demo-creds code{background:rgba(255,255,255,.7);padding:1px 6px;border-radius:4px;font-weight:600;color:var(--primary-darker);}
.login-foot{text-align:center;margin-top:18px;font-size:11.5px;color:var(--grey-dark);}

/* ============================================================
   APP SHELL (Admin + Employee)
   ============================================================ */
.app{display:none;min-height:100vh;}
.app.active{display:block;}
.app-layout{display:flex;min-height:100vh;}

/* Sidebar */
.sidebar{
  width:var(--sidebar-w);background:var(--white);border-right:1px solid var(--grey-light);
  display:flex;flex-direction:column;position:fixed;top:0;left:0;bottom:0;z-index:100;
  transition:transform var(--transition);
}
.sidebar-header{
  padding:22px 22px;display:flex;align-items:center;gap:12px;border-bottom:1px solid var(--grey-light);
}
.sidebar-header .logo-mark{width:42px;height:42px;border-radius:12px;font-size:16px;}
.sidebar-header .brand{display:flex;flex-direction:column;}
.sidebar-header .brand strong{font-size:15px;font-weight:800;color:var(--text);letter-spacing:.3px;}
.sidebar-header .brand span{font-size:10.5px;color:var(--text-light);text-transform:uppercase;letter-spacing:.5px;}
.sidebar-nav{flex:1;padding:14px 12px;overflow-y:auto;}
.nav-section-title{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:var(--grey-dark);font-weight:700;padding:14px 12px 6px;}
.nav-item{
  display:flex;align-items:center;gap:12px;padding:11px 14px;border-radius:var(--radius-sm);
  color:var(--text-light);font-weight:500;font-size:13.5px;transition:var(--transition);width:100%;text-align:left;
}
.nav-item:hover{background:var(--secondary);color:var(--primary-dark);}
.nav-item.active{background:var(--primary);color:#fff;box-shadow:0 4px 12px rgba(33,150,243,.25);}
.nav-item svg{width:19px;height:19px;flex-shrink:0;}
.nav-badge{margin-left:auto;background:var(--danger);color:#fff;font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;}
.nav-item.active .nav-badge{background:rgba(255,255,255,.25);}
.sidebar-footer{padding:14px 12px;border-top:1px solid var(--grey-light);}
.user-mini{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:var(--radius-sm);cursor:pointer;transition:var(--transition);}
.user-mini:hover{background:var(--bg);}
.user-mini .avatar{width:36px;height:36px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:13px;}
.user-mini .uinfo{flex:1;min-width:0;}
.user-mini .uinfo strong{display:block;font-size:13px;color:var(--text);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.user-mini .uinfo span{font-size:11px;color:var(--text-light);}

/* Main */
.main{flex:1;margin-left:var(--sidebar-w);display:flex;flex-direction:column;min-width:0;}
.topbar{
  background:var(--white);border-bottom:1px solid var(--grey-light);
  padding:14px 26px;display:flex;align-items:center;gap:18px;position:sticky;top:0;z-index:50;
}
.hamburger{display:none;width:38px;height:38px;border-radius:9px;background:var(--bg);align-items:center;justify-content:center;color:var(--text);}
.hamburger:hover{background:var(--grey-light);}
.topbar-title{display:flex;flex-direction:column;}
.topbar-title h2{font-size:18px;font-weight:700;color:var(--text);}
.topbar-title p{font-size:12px;color:var(--text-light);}
.topbar-spacer{flex:1;}
.clock-box{
  display:flex;align-items:center;gap:14px;background:var(--secondary);padding:9px 16px;border-radius:var(--radius-sm);
}
.clock-box .date{font-size:12px;color:var(--text-light);font-weight:500;}
.clock-box .time{font-size:18px;font-weight:700;color:var(--primary-dark);font-variant-numeric:tabular-nums;}
.topbar-right{display:flex;align-items:center;gap:10px;}
.icon-btn{width:40px;height:40px;border-radius:10px;background:var(--bg);display:flex;align-items:center;justify-content:center;color:var(--text-light);transition:var(--transition);}
.icon-btn:hover{background:var(--grey-light);color:var(--text);}

.content{padding:26px;flex:1;}
.view{display:none;animation:fade .3s ease;}
.view.active{display:block;}
@keyframes fade{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:translateY(0);}}

.page-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:14px;}
.page-head h3{font-size:20px;font-weight:700;color:var(--text);}
.page-head p{font-size:13px;color:var(--text-light);margin-top:2px;}
.page-actions{display:flex;gap:10px;flex-wrap:wrap;}

/* Summary cards */
.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px;}
.stat-card{
  background:var(--white);border-radius:var(--radius);padding:20px;box-shadow:var(--shadow);
  display:flex;align-items:flex-start;gap:14px;transition:var(--transition);position:relative;overflow:hidden;
}
.stat-card::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:var(--primary);}
.stat-card:hover{transform:translateY(-3px);box-shadow:var(--shadow-md);}
.stat-card .stat-icon{width:46px;height:46px;border-radius:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.stat-card .stat-icon svg{width:22px;height:22px;}
.stat-card .stat-body{flex:1;}
.stat-card .stat-label{font-size:11.5px;color:var(--text-light);font-weight:600;text-transform:uppercase;letter-spacing:.4px;}
.stat-card .stat-value{font-size:26px;font-weight:800;color:var(--text);margin-top:3px;font-variant-numeric:tabular-nums;}
.stat-card .stat-sub{font-size:11px;color:var(--text-light);margin-top:2px;}
.bg-primary{background:var(--secondary);color:var(--primary);}
.bg-accent{background:#E8F5E9;color:var(--accent);}
.bg-warning{background:var(--warning-light);color:var(--warning);}
.bg-danger{background:var(--danger-light);color:var(--danger);}
.bg-info{background:#E0F7FA;color:var(--info);}
.stat-card.accent-bar::before{background:var(--accent);}
.stat-card.warning-bar::before{background:var(--warning);}
.stat-card.danger-bar::before{background:var(--danger);}
.stat-card.info-bar::before{background:var(--info);}

/* Cards / panels */
.panel{
  background:var(--white);border-radius:var(--radius);box-shadow:var(--shadow);
  margin-bottom:20px;overflow:hidden;
}
.panel-head{padding:18px 22px;border-bottom:1px solid var(--grey-light);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;}
.panel-head h4{font-size:15px;font-weight:700;color:var(--text);display:flex;align-items:center;gap:8px;}
.panel-head h4 svg{width:18px;height:18px;color:var(--primary);}
.panel-body{padding:22px;}
.panel-body.flush{padding:0;}

.grid-2{display:grid;grid-template-columns:1.4fr 1fr;gap:20px;}
.grid-2-even{display:grid;grid-template-columns:1fr 1fr;gap:20px;}

/* Charts */
.chart-wrap{position:relative;height:280px;}
.chart-wrap.small{height:220px;}

/* Tables */
.table-toolbar{display:flex;gap:10px;flex-wrap:wrap;padding:16px 22px;border-bottom:1px solid var(--grey-light);align-items:center;}
.search-box{position:relative;flex:1;min-width:200px;}
.search-box .icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--grey-dark);width:16px;height:16px;}
.search-box input{width:100%;padding:10px 12px 10px 36px;border:1.5px solid var(--grey-light);border-radius:8px;background:var(--bg);transition:var(--transition);}
.search-box input:focus{outline:none;border-color:var(--primary);background:#fff;box-shadow:0 0 0 3px rgba(33,150,243,.1);}
.filter-select{padding:10px 12px;border:1.5px solid var(--grey-light);border-radius:8px;background:var(--bg);color:var(--text);font-weight:500;cursor:pointer;}
.filter-select:focus{outline:none;border-color:var(--primary);}
.table-scroll{overflow-x:auto;}
table.data-table{width:100%;border-collapse:collapse;font-size:13px;}
.data-table thead th{
  text-align:left;padding:13px 16px;background:var(--bg);color:var(--text-light);
  font-weight:600;font-size:11.5px;text-transform:uppercase;letter-spacing:.4px;
  border-bottom:1px solid var(--grey-light);white-space:nowrap;
}
.data-table tbody td{padding:13px 16px;border-bottom:1px solid var(--grey-light);color:var(--text);vertical-align:middle;}
.data-table tbody tr{transition:var(--transition);}
.data-table tbody tr:hover{background:var(--secondary);}
.data-table tbody tr:last-child td{border-bottom:none;}
.cell-strong{font-weight:600;}
.cell-muted{color:var(--text-light);}
.cell-id{font-family:'Inter',monospace;font-weight:600;color:var(--primary-dark);}
.avatar-cell{display:flex;align-items:center;gap:10px;}
.avatar{width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:13px;flex-shrink:0;object-fit:cover;}
.avatar img{width:100%;height:100%;border-radius:50%;object-fit:cover;}
.row-actions{display:flex;gap:6px;}
.row-actions button{width:30px;height:30px;border-radius:7px;background:var(--bg);display:flex;align-items:center;justify-content:center;color:var(--text-light);transition:var(--transition);}
.row-actions button:hover{background:var(--secondary);color:var(--primary);}
.row-actions button.del:hover{background:var(--danger-light);color:var(--danger);}

/* Badges */
.badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:11.5px;font-weight:600;}
.badge::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor;}
.badge-present{background:#E8F5E9;color:var(--accent-dark);}
.badge-late{background:var(--warning-light);color:var(--warning);}
.badge-absent{background:var(--danger-light);color:var(--danger);}
.badge-halfday{background:#FFE0B2;color:#E65100;}
.badge-leave{background:var(--secondary);color:var(--primary-dark);}
.badge-checkedout{background:#E0F7FA;color:#00838F;}

/* Employee cards grid */
.emp-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:16px;padding:22px;}
.emp-card{
  background:var(--white);border:1px solid var(--grey-light);border-radius:var(--radius);
  overflow:hidden;transition:var(--transition);
}
.emp-card:hover{transform:translateY(-4px);box-shadow:var(--shadow-md);border-color:var(--primary-light);}
.emp-card-head{height:70px;background:linear-gradient(135deg,var(--primary) 0%,var(--primary-dark) 100%);position:relative;}
.emp-card-body{padding:16px;text-align:center;margin-top:-34px;}
.emp-card-body .avatar{width:68px;height:68px;border-radius:50%;border:3px solid var(--white);margin:0 auto 10px;font-size:22px;box-shadow:var(--shadow);}
.emp-card-body .avatar img{border-radius:50%;}
.emp-card-body h4{font-size:15px;font-weight:700;color:var(--text);}
.emp-card-body .pos{font-size:12px;color:var(--text-light);margin-bottom:4px;}
.emp-card-body .dept{font-size:11px;color:var(--primary);font-weight:600;margin-bottom:12px;}
.emp-card-meta{display:flex;flex-direction:column;gap:6px;text-align:left;padding:12px 0;border-top:1px solid var(--grey-light);font-size:12px;color:var(--text-light);}
.emp-card-meta div{display:flex;align-items:center;gap:8px;}
.emp-card-meta svg{width:14px;height:14px;color:var(--grey-dark);}
.emp-card-actions{display:flex;gap:8px;padding-top:12px;border-top:1px solid var(--grey-light);}
.emp-card-actions button{flex:1;}

/* Modal */
.modal-overlay{
  position:fixed;inset:0;background:rgba(13,71,161,.32);backdrop-filter:blur(3px);
  display:none;align-items:center;justify-content:center;z-index:1000;padding:20px;
}
.modal-overlay.active{display:flex;animation:fade .2s ease;}
.modal{
  background:var(--white);border-radius:var(--radius);width:100%;max-width:560px;
  max-height:92vh;overflow-y:auto;box-shadow:var(--shadow-lg);animation:rise .3s ease;
}
.modal.lg{max-width:760px;}
.modal-head{padding:20px 24px;border-bottom:1px solid var(--grey-light);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:var(--white);z-index:2;}
.modal-head h4{font-size:17px;font-weight:700;color:var(--text);}
.modal-head .close{width:34px;height:34px;border-radius:8px;display:flex;align-items:center;justify-content:center;color:var(--text-light);background:var(--bg);}
.modal-head .close:hover{background:var(--danger-light);color:var(--danger);}
.modal-body{padding:24px;}
.modal-foot{padding:16px 24px;border-top:1px solid var(--grey-light);display:flex;justify-content:flex-end;gap:10px;position:sticky;bottom:0;background:var(--white);}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
.form-field{margin-bottom:16px;}
.form-field label{display:block;font-size:12px;font-weight:600;color:var(--text-light);margin-bottom:7px;}
.form-field label .req{color:var(--danger);}
.form-field input,.form-field select,.form-field textarea{
  width:100%;padding:11px 14px;border:1.5px solid var(--grey-light);border-radius:8px;background:var(--white);transition:var(--transition);color:var(--text);
}
.form-field input:focus,.form-field select:focus,.form-field textarea:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(33,150,243,.1);}
.form-field .hint{font-size:11px;color:var(--grey-dark);margin-top:5px;}
.photo-upload{display:flex;align-items:center;gap:16px;}
.photo-preview{width:72px;height:72px;border-radius:50%;background:var(--secondary);display:flex;align-items:center;justify-content:center;color:var(--primary);font-weight:700;font-size:22px;overflow:hidden;flex-shrink:0;border:2px solid var(--grey-light);}
.photo-preview img{width:100%;height:100%;object-fit:cover;}
.photo-upload input[type=file]{display:none;}
.photo-controls{flex:1;}
.photo-controls .hint{font-size:11px;color:var(--text-light);}

/* Employee dashboard specific */
.emp-hero{
  background:linear-gradient(135deg,var(--primary) 0%,var(--primary-darker) 100%);
  border-radius:var(--radius);padding:28px;color:#fff;box-shadow:0 8px 24px rgba(33,150,243,.25);
  display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;margin-bottom:22px;position:relative;overflow:hidden;
}
.emp-hero::after{content:"";position:absolute;right:-40px;top:-40px;width:200px;height:200px;border-radius:50%;background:rgba(255,255,255,.08);}
.emp-hero::before{content:"";position:absolute;right:60px;bottom:-60px;width:140px;height:140px;border-radius:50%;background:rgba(255,255,255,.06);}
.emp-hero .emp-info{display:flex;align-items:center;gap:16px;position:relative;z-index:1;}
.emp-hero .avatar{width:64px;height:64px;border:3px solid rgba(255,255,255,.4);font-size:22px;}
.emp-hero .emp-info h3{font-size:20px;font-weight:700;}
.emp-hero .emp-info p{font-size:13px;opacity:.9;}
.emp-hero .emp-info .id-pill{display:inline-block;margin-top:5px;font-size:11px;background:rgba(255,255,255,.2);padding:3px 10px;border-radius:20px;font-weight:600;}
.clock-hero{text-align:center;position:relative;z-index:1;}
.clock-hero .big-time{font-size:40px;font-weight:800;font-variant-numeric:tabular-nums;letter-spacing:1px;line-height:1;}
.clock-hero .big-date{font-size:13px;opacity:.9;margin-top:6px;}

.attend-actions{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:22px;}
.attend-btn{
  flex:1;min-width:200px;padding:22px;border-radius:var(--radius);display:flex;align-items:center;gap:16px;
  transition:var(--transition);color:#fff;font-weight:600;position:relative;overflow:hidden;
}
.attend-btn.checkin{background:linear-gradient(135deg,var(--accent) 0%,var(--accent-dark) 100%);box-shadow:0 6px 18px rgba(73,154,19,.3);}
.attend-btn.checkout{background:linear-gradient(135deg,#FB8C00 0%,#E65100 100%);box-shadow:0 6px 18px rgba(251,140,0,.3);}
.attend-btn:hover:not(:disabled){transform:translateY(-3px);}
.attend-btn:disabled{background:var(--grey-light)!important;color:var(--grey-dark)!important;box-shadow:none!important;cursor:not-allowed;}
.attend-btn .ab-icon{width:48px;height:48px;border-radius:12px;background:rgba(255,255,255,.2);display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.attend-btn .ab-icon svg{width:24px;height:24px;}
.attend-btn .ab-text strong{display:block;font-size:16px;}
.attend-btn .ab-text span{font-size:12px;opacity:.85;display:block;margin-top:2px;font-variant-numeric:tabular-nums;}

.today-status{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:14px;}
.status-tile{background:var(--bg);border-radius:var(--radius-sm);padding:16px;text-align:center;}
.status-tile .st-label{font-size:11px;color:var(--text-light);font-weight:600;text-transform:uppercase;letter-spacing:.4px;}
.status-tile .st-value{font-size:20px;font-weight:800;color:var(--text);margin-top:5px;font-variant-numeric:tabular-nums;}
.status-tile .st-value.accent{color:var(--accent);}
.status-tile .st-value.warning{color:var(--warning);}

/* Activity feed */
.activity-list{display:flex;flex-direction:column;}
.activity-item{display:flex;gap:14px;padding:13px 0;border-bottom:1px solid var(--grey-light);}
.activity-item:last-child{border-bottom:none;}
.activity-dot{width:34px;height:34px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.activity-dot svg{width:16px;height:16px;}
.activity-body{flex:1;}
.activity-body p{font-size:13px;color:var(--text);}
.activity-body p strong{font-weight:600;}
.activity-body .meta{font-size:11.5px;color:var(--text-light);margin-top:2px;}

/* Toast */
.toast-wrap{position:fixed;top:20px;right:20px;z-index:2000;display:flex;flex-direction:column;gap:10px;}
.toast{
  background:var(--white);border-radius:var(--radius-sm);padding:14px 18px;box-shadow:var(--shadow-lg);
  display:flex;align-items:center;gap:12px;min-width:280px;max-width:380px;border-left:4px solid var(--primary);
  animation:slideIn .3s ease;
}
@keyframes slideIn{from{opacity:0;transform:translateX(40px);}to{opacity:1;transform:translateX(0);}}
.toast.success{border-left-color:var(--accent);}
.toast.error{border-left-color:var(--danger);}
.toast.warning{border-left-color:var(--warning);}
.toast .t-icon{width:22px;height:22px;flex-shrink:0;}
.toast.success .t-icon{color:var(--accent);}
.toast.error .t-icon{color:var(--danger);}
.toast.warning .t-icon{color:var(--warning);}
.toast.info .t-icon{color:var(--primary);}
.toast .t-body{flex:1;}
.toast .t-body strong{display:block;font-size:13.5px;font-weight:600;color:var(--text);}
.toast .t-body span{font-size:12px;color:var(--text-light);}

/* Empty state */
.empty{text-align:center;padding:50px 20px;color:var(--text-light);}
.empty svg{width:48px;height:48px;color:var(--grey);margin-bottom:12px;}
.empty h4{font-size:15px;color:var(--text);margin-bottom:4px;}
.empty p{font-size:13px;}

/* Profile modal */
.profile-grid{display:grid;grid-template-columns:1fr 1.4fr;gap:24px;}
.profile-photo-wrap{text-align:center;}
.profile-photo-big{width:130px;height:130px;border-radius:50%;margin:0 auto 14px;background:var(--secondary);display:flex;align-items:center;justify-content:center;color:var(--primary);font-weight:700;font-size:40px;overflow:hidden;border:4px solid var(--secondary);}
.profile-photo-big img{width:100%;height:100%;object-fit:cover;}
.profile-detail-row{display:flex;justify-content:space-between;padding:11px 0;border-bottom:1px solid var(--grey-light);font-size:13px;}
.profile-detail-row:last-child{border-bottom:none;}
.profile-detail-row .pdr-label{color:var(--text-light);font-weight:500;}
.profile-detail-row .pdr-value{color:var(--text);font-weight:600;}

/* Export buttons */
.export-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;}
.export-card{background:var(--bg);border-radius:var(--radius);padding:22px;text-align:center;transition:var(--transition);cursor:pointer;border:1.5px solid transparent;}
.export-card:hover{background:var(--white);border-color:var(--primary);box-shadow:var(--shadow);}
.export-card .ec-icon{width:54px;height:54px;border-radius:14px;display:flex;align-items:center;justify-content:center;margin:0 auto 12px;}
.export-card .ec-icon.green{background:#E8F5E9;color:var(--accent);}
.export-card .ec-icon.red{background:var(--danger-light);color:var(--danger);}
.export-card h4{font-size:14px;font-weight:700;color:var(--text);margin-bottom:4px;}
.export-card p{font-size:12px;color:var(--text-light);}

/* Mobile */
.backdrop{position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:99;display:none;}
.backdrop.active{display:block;}
@media(max-width:992px){
  .grid-2,.grid-2-even{grid-template-columns:1fr;}
  .profile-grid{grid-template-columns:1fr;}
}
@media(max-width:768px){
  :root{--sidebar-w:260px;}
  .sidebar{transform:translateX(-100%);}
  .sidebar.open{transform:translateX(0);}
  .main{margin-left:0;}
  .hamburger{display:flex;}
  .topbar{padding:12px 16px;}
  .content{padding:16px;}
  .clock-box .date{display:none;}
  .clock-box .time{font-size:15px;}
  .form-row{grid-template-columns:1fr;}
  .emp-hero{flex-direction:column;text-align:center;}
  .clock-hero .big-time{font-size:32px;}
  .topbar-title h2{font-size:15px;}
}
@media(max-width:480px){
  .attend-btn{min-width:100%;}
  .modal-body{padding:18px;}
}
</style>
</head>
<body>

<!-- ============================================================
     LOGIN SCREEN
     ============================================================ -->
<div id="loginScreen">
  <div class="login-card">
    <div class="login-logo">
      <div class="logo-mark">Y</div>
      <h1>YUVAPACK</h1>
      <p>Employee Attendance Management System</p>
    </div>

    <div class="login-tabs">
      <button class="login-tab active" data-tab="admin" onclick="switchLoginTab('admin')">Admin Login</button>
      <button class="login-tab" data-tab="employee" onclick="switchLoginTab('employee')">Employee Login</button>
    </div>

    <!-- Admin Login -->
    <form id="adminLoginForm" onsubmit="return adminLogin(event)">
      <div class="form-group">
        <label>Admin ID</label>
        <div class="input-wrap">
          <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
          <input class="form-input" id="adminId" placeholder="Enter Admin ID" autocomplete="off" required>
        </div>
      </div>
      <div class="form-group">
        <label>Password</label>
        <div class="input-wrap">
          <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
          <input class="form-input" id="adminPass" type="password" placeholder="Enter password" required>
        </div>
      </div>
      <button class="btn btn-primary" type="submit">Login as Admin</button>
    </form>

    <!-- Employee Login -->
    <form id="employeeLoginForm" style="display:none" onsubmit="return employeeLogin(event)">
      <div class="form-group">
        <label>Employee ID</label>
        <div class="input-wrap">
          <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M7 8h10M7 12h10M7 16h6"/></svg>
          <input class="form-input" id="empId" placeholder="Enter Employee ID" autocomplete="off" required>
        </div>
      </div>
      <div class="form-group">
        <label>Password</label>
        <div class="input-wrap">
          <svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
          <input class="form-input" id="empPass" type="password" placeholder="Enter password" required>
        </div>
      </div>
      <button class="btn btn-primary" type="submit">Login as Employee</button>
    </form>

    <p class="login-foot">yuvapack_attendance login screen</p>
  </div>
</div>

<!-- ============================================================
     ADMIN APP
     ============================================================ -->
<div id="adminApp" class="app">
  <div class="app-layout">
    <div class="backdrop" id="backdrop" onclick="closeSidebar()"></div>
    <aside class="sidebar" id="sidebar">
      <div class="sidebar-header">
        <div class="logo-mark">Y</div>
        <div class="brand"><strong>YUVAPACK</strong><span>Admin Panel</span></div>
      </div>
      <nav class="sidebar-nav">
        <div class="nav-section-title">Main</div>
        <button class="nav-item active" data-view="dashboard" onclick="showAdminView('dashboard')">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
          Dashboard
        </button>
        <button class="nav-item" data-view="employees" onclick="showAdminView('employees')">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
          Employees
        </button>
        <button class="nav-item" data-view="attendance" onclick="showAdminView('attendance')">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
          Attendance
        </button>
        <button class="nav-item" data-view="reports" onclick="showAdminView('reports')">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="13" y2="17"/></svg>
          Reports &amp; Export
        </button>
        <div class="nav-section-title">Account</div>
        <button class="nav-item" onclick="logout()">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
          Logout
        </button>
      </nav>
      <div class="sidebar-footer">
        <div class="user-mini">
          <div class="avatar" style="background:var(--primary)">A</div>
          <div class="uinfo"><strong>Administrator</strong><span>admin@yuvapack</span></div>
        </div>
      </div>
    </aside>

    <div class="main">
      <header class="topbar">
        <button class="hamburger" onclick="openSidebar()"><svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg></button>
        <div class="topbar-title">
          <h2 id="adminPageTitle">Dashboard</h2>
          <p id="adminPageSub">Overview of attendance &amp; workforce</p>
        </div>
        <div class="topbar-spacer"></div>
        <div class="clock-box">
          <div><div class="date" id="adminDate">—</div><div class="time" id="adminTime">--:--:--</div></div>
        </div>
        <div class="topbar-right">
          <button class="icon-btn" onclick="logout()" title="Logout"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg></button>
        </div>
      </header>

      <main class="content">
        <!-- DASHBOARD VIEW -->
        <section class="view active" id="view-dashboard">
          <div class="stats-grid" id="statsGrid"></div>
          <div class="grid-2">
            <div class="panel">
              <div class="panel-head"><h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>Attendance Overview</h4>
                <select class="filter-select" id="chartRange" onchange="renderChart()"><option value="today">Today</option><option value="week">This Week</option></select>
              </div>
              <div class="panel-body"><div class="chart-wrap"><canvas id="attChart"></canvas></div></div>
            </div>
            <div class="panel">
              <div class="panel-head"><h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>Recent Activity</h4></div>
              <div class="panel-body"><div class="activity-list" id="activityList"></div></div>
            </div>
          </div>
          <div class="panel">
            <div class="panel-head"><h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/></svg>Today's Attendance</h4>
              <div class="search-box" style="min-width:200px;max-width:260px"><svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg><input id="dashSearch" placeholder="Search employee..." oninput="renderDashTable()"></div>
            </div>
            <div class="table-scroll"><table class="data-table"><thead><tr><th>Employee</th><th>ID</th><th>Department</th><th>Check-In</th><th>Check-Out</th><th>Hours</th><th>Status</th></tr></thead><tbody id="dashTableBody"></tbody></table></div>
          </div>
        </section>

        <!-- EMPLOYEES VIEW -->
        <section class="view" id="view-employees">
          <div class="panel">
            <div class="table-toolbar">
              <div class="search-box"><svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg><input id="empSearch" placeholder="Search by name, ID, department..." oninput="renderEmployees()"></div>
              <select class="filter-select" id="empDeptFilter" onchange="renderEmployees()"><option value="">All Departments</option></select>
              <button class="btn btn-primary btn-sm" onclick="openEmpModal()"><svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>Add Employee</button>
            </div>
            <div id="empGridWrap"></div>
          </div>
        </section>

        <!-- ATTENDANCE VIEW -->
        <section class="view" id="view-attendance">
          <div class="panel">
            <div class="table-toolbar">
              <div class="search-box"><svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg><input id="attSearch" placeholder="Search employee or ID..." oninput="renderAttendanceTable()"></div>
              <input type="date" class="filter-select" id="attDateFilter" onchange="renderAttendanceTable()">
              <select class="filter-select" id="attStatusFilter" onchange="renderAttendanceTable()"><option value="">All Status</option><option>Present</option><option>Late</option><option>Absent</option><option>Half Day</option><option>On Leave</option></select>
              <button class="btn btn-accent btn-sm" onclick="markAbsentToday()">Mark Absent (No Check-in)</button>
            </div>
            <div class="table-scroll"><table class="data-table"><thead><tr><th>Employee</th><th>ID</th><th>Date</th><th>Check-In</th><th>Check-Out</th><th>Working Hrs</th><th>Overtime</th><th>Status</th></tr></thead><tbody id="attTableBody"></tbody></table></div>
          </div>
        </section>

        <!-- REPORTS VIEW -->
        <section class="view" id="view-reports">
          <div class="panel">
            <div class="panel-head"><h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>Export Reports</h4></div>
            <div class="panel-body">
              <p style="color:var(--text-light);margin-bottom:18px;">Download attendance and employee data in Excel or PDF format. Reports include all records currently stored in this browser.</p>
              <div class="export-grid">
                <div class="export-card" onclick="exportAttendanceExcel()">
                  <div class="ec-icon green"><svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="9" y1="15" x2="15" y2="15"/></svg></div>
                  <h4>Attendance → Excel</h4><p>All attendance records (.xlsx)</p>
                </div>
                <div class="export-card" onclick="exportEmployeesExcel()">
                  <div class="ec-icon green"><svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg></div>
                  <h4>Employees → Excel</h4><p>Employee directory (.xlsx)</p>
                </div>
                <div class="export-card" onclick="exportAttendancePDF()">
                  <div class="ec-icon red"><svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg></div>
                  <h4>Attendance → PDF</h4><p>Printable report (.pdf)</p>
                </div>
                <div class="export-card" onclick="exportSummaryPDF()">
                  <div class="ec-icon red"><svg viewBox="0 0 24 24" width="26" height="26" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg></div>
                  <h4>Summary → PDF</h4><p>Today's dashboard summary</p>
                </div>
              </div>
              <div style="margin-top:22px;padding:16px;background:var(--secondary);border-radius:var(--radius-sm);font-size:12.5px;color:var(--text-light);">
                <strong style="color:var(--primary-dark);">Storage info:</strong> Export Attendance Excel and Pdf Format Download Now.
              </div>
            </div>
          </div>
        </section>
      </main>
    </div>
  </div>
</div>

<!-- ============================================================
     EMPLOYEE APP
     ============================================================ -->
<div id="empApp" class="app">
  <div class="app-layout">
    <aside class="sidebar">
      <div class="sidebar-header">
        <div class="logo-mark">Y</div>
        <div class="brand"><strong>YUVAPACK</strong><span>Employee Portal</span></div>
      </div>
      <nav class="sidebar-nav">
        <div class="nav-section-title">Workspace</div>
        <button class="nav-item active">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
          My Attendance
        </button>
        <button class="nav-item" onclick="renderEmpHistory(); document.getElementById('empHistoryPanel').scrollIntoView({behavior:'smooth'})">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
          History
        </button>
        <div class="nav-section-title">Account</div>
        <button class="nav-item" onclick="logout()">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
          Logout
        </button>
      </nav>
      <div class="sidebar-footer">
        <div class="user-mini" id="empUserMini"></div>
      </div>
    </aside>

    <div class="main">
      <header class="topbar">
        <div class="topbar-title">
          <h2>My Attendance</h2>
          <p>Mark your check-in and check-out for today</p>
        </div>
        <div class="topbar-spacer"></div>
        <div class="clock-box">
          <div><div class="date" id="empDate">—</div><div class="time" id="empTime">--:--:--</div></div>
        </div>
        <div class="topbar-right">
          <button class="icon-btn" onclick="logout()" title="Logout"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg></button>
        </div>
      </header>

      <main class="content">
        <div class="emp-hero" id="empHero"></div>

        <div class="attend-actions">
          <button class="attend-btn checkin" id="checkinBtn" onclick="empCheckIn()">
            <div class="ab-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg></div>
            <div class="ab-text"><strong>Check In</strong><span id="checkinTime">Not checked in</span></div>
          </button>
          <button class="attend-btn checkout" id="checkoutBtn" onclick="empCheckOut()" disabled>
            <div class="ab-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg></div>
            <div class="ab-text"><strong>Check Out</strong><span id="checkoutTime">Check in first</span></div>
          </button>
        </div>

        <div class="panel">
          <div class="panel-head"><h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v20M2 12h20"/></svg>Today's Summary</h4></div>
          <div class="panel-body"><div class="today-status" id="todayStatus"></div></div>
        </div>

        <div class="panel" id="empHistoryPanel">
          <div class="panel-head"><h4><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/></svg>My Attendance History</h4></div>
          <div class="table-scroll"><table class="data-table"><thead><tr><th>Date</th><th>Check-In</th><th>Check-Out</th><th>Working Hrs</th><th>Overtime</th><th>Status</th></tr></thead><tbody id="empHistoryBody"></tbody></table></div>
        </div>
      </main>
    </div>
  </div>
</div>

<!-- ============================================================
     MODALS & TOAST
     ============================================================ -->
<div class="modal-overlay" id="empModalOverlay">
  <div class="modal" id="empModal"></div>
</div>
<div class="modal-overlay" id="profileModalOverlay">
  <div class="modal" id="profileModal"></div>
</div>
<div class="toast-wrap" id="toastWrap"></div>

<script>
/* ============================================================
   YUVAPACK — Application Logic
   Pure vanilla JS · LocalStorage persistence
   ============================================================ */

/* ---------- Constants ---------- */
const OFFICE_START_MIN = 9*60;        // 09:00 AM
const GRACE_MIN = 15;                  // 15 minutes
const LATE_AFTER_MIN = OFFICE_START_MIN + GRACE_MIN;  // 09:15 AM
const HALF_DAY_MIN = 12*60;           // 12:00 PM
const STANDARD_HOURS = 8;

const LS_KEYS = {
  employees:'yuvapack_employees',
  attendance:'yuvapack_attendance',
  session:'yuvapack_session',
  seeded:'yuvapack_seeded_v1'
};

const ADMIN_CRED = { id:'admin', pass:'admin123' };
const DEFAULT_EMP_PASS = 'emp123';

const DEPARTMENTS = ['Designer','Production','Accounts','HR','Marketing','Employee','Quality','Machine Operator',];

/* ---------- Data layer ---------- */
const DB = {
  getEmployees(){
    try{ return JSON.parse(localStorage.getItem(LS_KEYS.employees))||[]; }catch(e){ return []; }
  },
  saveEmployees(arr){ localStorage.setItem(LS_KEYS.employees, JSON.stringify(arr)); },
  getAttendance(){
    try{ return JSON.parse(localStorage.getItem(LS_KEYS.attendance))||[]; }catch(e){ return []; }
  },
  saveAttendance(arr){ localStorage.setItem(LS_KEYS.attendance, JSON.stringify(arr)); },
  getSession(){
    try{ return JSON.parse(localStorage.getItem(LS_KEYS.session))||null; }catch(e){ return null; }
  },
  setSession(s){ localStorage.setItem(LS_KEYS.session, JSON.stringify(s)); },
  clearSession(){ localStorage.removeItem(LS_KEYS.session); }
};

/* ---------- Seed demo data ---------- */
function seedData(){
  if(localStorage.getItem(LS_KEYS.seeded)) return;

  const palette = ['#2196F3','#499A13','#FB8C00','#00ACC1','#E53935','#5E35B1','#00897B'];
  const emps = [

  ];
  DB.saveEmployees(emps);

  // Seed attendance for the last 7 days (excluding today, which starts fresh)
  const att = [];
  const today = new Date();
  for(let d=7; d>=1; d--){
    const day = new Date(today);
    day.setDate(today.getDate()-d);
    if(day.getDay()===0) continue; // skip Sundays
    const ds = dateStr(day);
    emps.forEach((e,i)=>{
      // varied patterns
      const pattern = (i + d) % 5;
      if(pattern===4){ att.push({ id:recId(), employeeId:e.id, date:ds, checkIn:null, checkOut:null, workingHours:0, overtimeHours:0, status:'On Leave' }); return; }
      let ciH=9, ciM=5 + (i*3+d*2)%50;
      if(pattern===1){ ciH=9; ciM=25+(i*2)%30; } // late
      if(pattern===2 && d%2===0){ ciH=12; ciM=10+(i*2)%20; } // half day
      const ci = new Date(day); ci.setHours(ciH, ciM, 0, 0);
      const co = new Date(day); co.setHours(17 + (i%2), (i*7)%60, 0, 0);
      const wh = (co - ci)/3600000;
      att.push({
        id:recId(), employeeId:e.id, date:ds,
        checkIn:formatTime(ci), checkOut:formatTime(co),
        workingHours:round2(wh), overtimeHours:round2(Math.max(0, wh-STANDARD_HOURS)),
        status:calcStatusFromTime(formatTime(ci))
      });
    });
  }
  DB.saveAttendance(att);
  localStorage.setItem(LS_KEYS.seeded,'1');
}

function recId(){ return 'rec_'+Date.now().toString(36)+Math.random().toString(36).slice(2,7); }

/* ---------- Date / time helpers ---------- */
function dateStr(d){ // YYYY-MM-DD local
  const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), dd=String(d.getDate()).padStart(2,'0');
  return `${y}-${m}-${dd}`;
}
function formatTime(d){
  let h=d.getHours(), m=d.getMinutes();
  const ap = h>=12?'PM':'AM';
  h = h%12; if(h===0) h=12;
  return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')} ${ap}`;
}
function timeToMinutes(t){ // "09:05 AM" -> minutes
  if(!t) return null;
  const m = t.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
  if(!m) return null;
  let h=parseInt(m[1],10), min=parseInt(m[2],10);
  const ap=m[3].toUpperCase();
  if(ap==='PM'&&h!==12) h+=12;
  if(ap==='AM'&&h===12) h=0;
  return h*60+min;
}
function calcStatusFromTime(checkInStr){
  if(!checkInStr) return 'Absent';
  const m = timeToMinutes(checkInStr);
  if(m<=LATE_AFTER_MIN) return 'Present';
  if(m<=HALF_DAY_MIN) return 'Late';
  return 'Half Day';
}
function round2(n){ return Math.round(n*100)/100; }

function todayDateStr(){ return dateStr(new Date()); }

function prettyDate(d){
  const dt = (d instanceof Date)? d : new Date(d+'T00:00:00');
  return dt.toLocaleDateString('en-US',{weekday:'short',year:'numeric',month:'short',day:'numeric'});
}

/* ---------- Toast ---------- */
function toast(title, msg, type='info'){
  const wrap = document.getElementById('toastWrap');
  const icons = {
    success:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
    error:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
    warning:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    info:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>'
  };
  const el = document.createElement('div');
  el.className = 'toast '+type;
  el.innerHTML = `<span class="t-icon">${icons[type]||icons.info}</span><div class="t-body"><strong>${title}</strong><span>${msg}</span></div>`;
  wrap.appendChild(el);
  setTimeout(()=>{ el.style.transition='opacity .3s,transform .3s'; el.style.opacity='0'; el.style.transform='translateX(40px)'; setTimeout(()=>el.remove(),300); }, 3600);
}

/* ---------- Login ---------- */
function switchLoginTab(tab){
  document.querySelectorAll('.login-tab').forEach(t=>t.classList.toggle('active', t.dataset.tab===tab));
  document.getElementById('adminLoginForm').style.display = tab==='admin'?'block':'none';
  document.getElementById('employeeLoginForm').style.display = tab==='employee'?'block':'none';
}


async function adminLogin(e) {
    e.preventDefault();

    const id = document.getElementById('adminId').value.trim();
    const pass = document.getElementById('adminPass').value;

    try {
        const response = await fetch('login_api.php', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            credentials: 'same-origin',
            body: JSON.stringify({
                username: id,
                password: pass
            })
        });

        const result = await response.json();

        if (result.success && result.user.role === 'admin') {

            DB.setSession({
                role: 'admin',
                id: result.user.username
            });

            showApp('admin');

            toast(
                'Welcome back',
                'Logged in as Administrator',
                'success'
            );

        } else {
            toast(
                'Login failed',
                result.message || 'Invalid admin credentials',
                'error'
            );
        }

    } catch (error) {
        console.error(error);

        toast(
            'Connection error',
            'Unable to connect to server',
            'error'
        );
    }

    return false;
}



async function employeeLogin(e) {
  e.preventDefault();

  const id = document.getElementById('empId').value.trim();
  const pass = document.getElementById('empPass').value;

  try {
    const response = await fetch('./employee_login_api.php', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      credentials: 'same-origin',
      body: JSON.stringify({
        employee_code: id,
        password: pass
      })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      toast(
        'Login failed',
        result.message || 'Invalid employee ID or password',
        'error'
      );
      return false;
    }

    const emp = result.employee;

    _currentEmp = {
      id: emp.employee_code,
      name: emp.name,
      position: emp.position || '',
      department: emp.department || '',
      phone: emp.phone || '',
      joiningDate: emp.joining_date || '',
      photo: '',
      color: '#2196F3'
    };

    DB.setSession({
      role: 'employee',
      id: emp.employee_code
    });

    showApp('employee');

try {
  await syncTodayAttendance();
  updateEmpAttendanceUI();
} catch (error) {
  console.error('Attendance loading error:', error);
}

toast(
  'Welcome ' + emp.name.split(' ')[0],
  'You are now logged in',
  'success'
);

  } catch (error) {
    console.error('Employee Login Error:', error);
    toast(
      'Server Error',
      'Cannot connect to employee login API',
      'error'
    );
  }

  return false;
}


function logout(){
  DB.clearSession();
  document.getElementById('adminApp').classList.remove('active');
  document.getElementById('empApp').classList.remove('active');
  document.getElementById('loginScreen').style.display='flex';
  // reset forms
  document.getElementById('adminId').value='';
  document.getElementById('adminPass').value='';
  document.getElementById('empId').value='';
  document.getElementById('empPass').value='';
  toast('Logged out','Session ended successfully','info');
}

function showApp(role){
  document.getElementById('loginScreen').style.display='none';
  if(role==='admin'){
    document.getElementById('adminApp').classList.add('active');
    showAdminView('dashboard');
  } else {
    document.getElementById('empApp').classList.add('active');
    initEmployeeApp();
  }
}

/* ---------- Live clock ---------- */
let clockTimer=null;
function startClock(){
  if(clockTimer) clearInterval(clockTimer);
  const tick=()=>{
    const now=new Date();
    const dStr = now.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric',year:'numeric'});
    const tStr = now.toLocaleTimeString('en-US',{hour12:true,hour:'2-digit',minute:'2-digit',second:'2-digit'});
    const ad=document.getElementById('adminDate'), at=document.getElementById('adminTime');
    const ed=document.getElementById('empDate'), et=document.getElementById('empTime');
    if(ad) ad.textContent=dStr; if(at) at.textContent=tStr;
    if(ed) ed.textContent=dStr; if(et) et.textContent=tStr;
    const hero=document.querySelector('.clock-hero .big-time');
    const heroD=document.querySelector('.clock-hero .big-date');
    if(hero) hero.textContent=tStr;
    if(heroD) heroD.textContent=dStr;
  };
  tick();
  clockTimer=setInterval(tick,1000);
}

/* ---------- Admin views ---------- */
const ADMIN_TITLES = {
  dashboard:['Dashboard','Overview of attendance &amp; workforce'],
  employees:['Employees','Manage your workforce directory'],
  attendance:['Attendance Records','Track and review attendance'],
  reports:['Reports &amp; Export','Download data in Excel or PDF']
};
function showAdminView(v){
  document.querySelectorAll('#adminApp .view').forEach(s=>s.classList.remove('active'));
  document.getElementById('view-'+v).classList.add('active');
  document.querySelectorAll('#adminApp .nav-item').forEach(n=>n.classList.toggle('active', n.dataset.view===v));
  document.getElementById('adminPageTitle').innerHTML = ADMIN_TITLES[v][0];
  document.getElementById('adminPageSub').innerHTML = ADMIN_TITLES[v][1];
  closeSidebar();
  if(v==='dashboard'){ renderDashboard(); }
  if(v==='employees'){ populateDeptFilter(); renderEmployees(); }
  if(v==='attendance'){ renderAttendanceTable(); }
}

/* ---------- Dashboard render ---------- */

async function renderDashboard() {
  try {
    const response = await fetch('./admin_dashboard_api.php', {
      method: 'GET',
      credentials: 'same-origin'
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.message || 'Dashboard loading failed');
    }

    // Save MySQL data in dashboard cache
    _dashboardEmployees = result.employees.map(e => ({
      id: e.employee_code,
      name: e.name,
      position: e.position || '',
      department: e.department || '',
      phone: e.phone || '',
      joiningDate: e.joining_date || '',
      photo: '',
      color: '#2196F3'
    }));

    _dashboardTodayAttendance = result.todayAttendance.map(a => ({
      id: a.id,
      employeeId: a.employee_code,
      date: a.attendance_date,
      checkIn: a.check_in,
      checkOut: a.check_out,
      workingHours: Number(a.total_hours || 0),
      overtimeHours: Math.max(
        0,
        round2(Number(a.total_hours || 0) - STANDARD_HOURS)
      ),
      status: a.status
    }));

    _dashboardRecentAttendance = result.recentActivity;

    const emps = _dashboardEmployees;
    const todayAtt = _dashboardTodayAttendance;

    let present = 0, late = 0, absent = 0;
    let halfday = 0, leave = 0;

    const recordedIds = new Set();

    todayAtt.forEach(a => {
      recordedIds.add(a.employeeId);

      if (a.status === 'Present') present++;
      else if (a.status === 'Late') late++;
      else if (a.status === 'Half Day') halfday++;
      else if (a.status === 'On Leave') leave++;
      else if (a.status === 'Absent') absent++;
    });

    absent += emps.length - recordedIds.size;

    const totalHours = todayAtt.reduce(
      (sum, a) => sum + a.workingHours,
      0
    );

    const stats = [
      {
        label: 'Total Employees',
        value: emps.length,
        sub: 'Active workforce',
        icon: empIcon(),
        cls: 'bg-primary',
        bar: ''
      },
      {
        label: 'Present Today',
        value: present,
        sub: 'On time',
        icon: checkIcon(),
        cls: 'bg-accent',
        bar: 'accent-bar'
      },
      {
        label: 'Late Employees',
        value: late,
        sub: 'After 09:15 AM',
        icon: clockIcon(),
        cls: 'bg-warning',
        bar: 'warning-bar'
      },
      {
        label: 'Absent Today',
        value: absent,
        sub: 'No check-in',
        icon: xIcon(),
        cls: 'bg-danger',
        bar: 'danger-bar'
      },
      {
        label: 'On Leave',
        value: leave,
        sub: 'Approved leave',
        icon: planeIcon(),
        cls: 'bg-info',
        bar: 'info-bar'
      },
      {
        label: 'Total Working Hrs',
        value: round2(totalHours),
        sub: 'Today combined',
        icon: hourIcon(),
        cls: 'bg-primary',
        bar: ''
      }
    ];

    document.getElementById('statsGrid').innerHTML =
      stats.map(s => `
        <div class="stat-card ${s.bar}">
          <div class="stat-icon ${s.cls}">${s.icon}</div>
          <div class="stat-body">
            <div class="stat-label">${s.label}</div>
            <div class="stat-value">${s.value}</div>
            <div class="stat-sub">${s.sub}</div>
          </div>
        </div>
      `).join('');

    renderChart();
    renderActivity();
    renderDashTable();

  } catch (error) {
    console.error('Dashboard Error:', error);   
    toast('Dashboard Error', error.message, 'error');
  }
}


  function renderChart(){
    const ctx = document.getElementById('attChart');
    if(!ctx) return;
    const range = document.getElementById('chartRange').value;
    const att = DB.getAttendance();
    let records;
    if(range==='today'){
      records = att.filter(a=>a.date===todayDateStr());
    } else {
      const start = new Date(); start.setDate(start.getDate()-6);
      const sd = dateStr(start);
      records = att.filter(a=>a.date>=sd);
    }
    const counts = {Present:0,Late:0,Absent:0,'Half Day':0,'On Leave':0};
    records.forEach(a=>{ if(counts[a.status]!==undefined) counts[a.status]++; });
    if(window._attChart) window._attChart.destroy();
    window._attChart = new Chart(ctx,{
      type:'doughnut',
      data:{
        labels:['Present','Late','Absent','Half Day','On Leave'],
        datasets:[{
          data:[counts.Present,counts.Late,counts.Absent,counts['Half Day'],counts['On Leave']],
          backgroundColor:['#499A13','#FB8C00','#E53935','#FFB74D','#2196F3'],
          borderWidth:0, hoverOffset:8
        }]
      },
      options:{
        responsive:true, maintainAspectRatio:false,
        cutout:'62%',
        plugins:{ legend:{ position:'right', labels:{ usePointStyle:true, pointStyle:'circle', padding:14, font:{size:12, family:'Inter'} } },
          tooltip:{ callbacks:{ label:c=>` ${c.label}: ${c.parsed}` } } }
      }
    });
  }


function renderActivity() {
  const sorted = _dashboardRecentAttendance;
  const list = document.getElementById('activityList');

  if (sorted.length === 0) {
    list.innerHTML = `
      <div class="empty">
        <h4>No activity yet</h4>
        <p>Employee check-ins will appear here.</p>
      </div>
    `;
    return;
  }

  const map = {
    'Present': ['bg-accent', checkIcon()],
    'Late': ['bg-warning', clockIcon()],
    'Half Day': ['bg-warning', hourIcon()],
    'Absent': ['bg-danger', xIcon()],
    'On Leave': ['bg-info', planeIcon()]
  };

  list.innerHTML = sorted.map(a => {
    const [ic, c] = map[a.status] || ['bg-primary', checkIcon()];

    return `
      <div class="activity-item">
        <div class="activity-dot ${ic}" style="background:var(--secondary)">
          ${c}
        </div>
        <div class="activity-body">
          <p>
            <strong>${a.name}</strong>
            marked ${String(a.status).toLowerCase()}
          </p>
          <div class="meta">
            ${a.check_in ? 'Check-in ' + a.check_in : 'No check-in recorded'}
            · ${prettyDate(a.attendance_date)}
          </div>
        </div>
      </div>
    `;
  }).join('');
}



function renderDashTable() {
  const q = (
    document.getElementById('dashSearch').value || ''
  ).toLowerCase();

  const att = _dashboardTodayAttendance;
  const emps = _dashboardEmployees;

  const rows = emps.map(e => {
    const a = att.find(x => x.employeeId === e.id);
    return { emp: e, att: a };
  }).filter(r => {
    const e = r.emp;

    return (
      !q ||
      e.name.toLowerCase().includes(q) ||
      e.id.toLowerCase().includes(q) ||
      e.department.toLowerCase().includes(q)
    );
  });

  const body = document.getElementById('dashTableBody');

  if (rows.length === 0) {
    body.innerHTML = `
      <tr>
        <td colspan="7">
          <div class="empty">
            <h4>No employees</h4>
            <p>Add employees from the Employees tab.</p>
          </div>
        </td>
      </tr>
    `;
    return;
  }

  body.innerHTML = rows.map(r => {
    const a = r.att;
    const status = a ? a.status : 'Absent';

    return `
      <tr>
        <td>
          <div class="avatar-cell">
            ${avatarHtml(r.emp)}
            <div>
              <div class="cell-strong">${r.emp.name}</div>
              <div class="cell-muted">${r.emp.position}</div>
            </div>
          </div>
        </td>
        <td class="cell-id">${r.emp.id}</td>
        <td>${r.emp.department}</td>
        <td>${a && a.checkIn ? a.checkIn : '<span class="cell-muted">—</span>'}</td>
        <td>${a && a.checkOut ? a.checkOut : '<span class="cell-muted">—</span>'}</td>
        <td>${a && a.checkOut ? (a.workingHours + 'h') : '<span class="cell-muted">—</span>'}</td>
        <td>${badgeHtml(status)}</td>
      </tr>
    `;
  }).join('');
}


/* ---------- Employees management ---------- */
function populateDeptFilter(){
  const sel = document.getElementById('empDeptFilter');
  const cur = sel.value;
  const depts = [...new Set(DB.getEmployees().map(e=>e.department))].sort();
  sel.innerHTML = '<option value="">All Departments</option>' + depts.map(d=>`<option>${d}</option>`).join('');
  sel.value = cur;
}

function renderEmployees(){
  const q=(document.getElementById('empSearch').value||'').toLowerCase();
  const dept=document.getElementById('empDeptFilter').value;
  let emps = DB.getEmployees();
  if(q) emps=emps.filter(e=>e.name.toLowerCase().includes(q)||e.id.toLowerCase().includes(q)||e.position.toLowerCase().includes(q)||e.department.toLowerCase().includes(q));
  if(dept) emps=emps.filter(e=>e.department===dept);
  const wrap=document.getElementById('empGridWrap');
  if(emps.length===0){
    wrap.innerHTML='<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg><h4>No employees found</h4><p>Try adjusting your search or add a new employee.</p></div>';
    return;
  }
  wrap.innerHTML = '<div class="emp-grid">'+emps.map(e=>`
    <div class="emp-card">
      <div class="emp-card-head"></div>
      <div class="emp-card-body">
        ${avatarHtml(e,'avatar',68)}
        <h4>${e.name}</h4>
        <div class="pos">${e.position}</div>
        <div class="dept">${e.department}</div>
        <div class="emp-card-meta">
          <div><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="16" rx="2"/></svg> ${e.id}</div>
          <div><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92z"/></svg> ${e.phone||'—'}</div>
          <div><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg> Joined ${prettyDate(e.joiningDate)}</div>
        </div>
        <div class="emp-card-actions">
          <button class="btn btn-ghost btn-sm" onclick="viewEmployee('${e.id}')">View</button>
          <button class="btn btn-outline btn-sm" onclick="openEmpModal('${e.id}')">Edit</button>
          <button class="btn btn-danger btn-sm" onclick="deleteEmployee('${e.id}')">Delete</button>
        </div>
      </div>
    </div>`).join('')+'</div>';
}

/* Employee modal (add/edit) */
let _editingEmpId = null;
let _pendingPhoto = null;

function openEmpModal(id){
  _editingEmpId = id||null;
  _pendingPhoto = null;
  const emp = id? DB.getEmployees().find(e=>e.id===id) : null;
  const isEdit = !!emp;
  const initials = emp? emp.name.split(' ').map(s=>s[0]).join('').slice(0,2).toUpperCase() : 'YA';
  const photoHtml = (emp && emp.photo) ? `<img src="${emp.photo}" alt="">` : initials;

  document.getElementById('empModal').innerHTML = `
    <div class="modal-head"><h4>${isEdit?'Edit Employee':'Add New Employee'}</h4><button class="close" onclick="closeEmpModal()"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button></div>
    <div class="modal-body">
      <form id="empForm" onsubmit="return saveEmployee(event)">
        <div class="form-field">
          <label>Profile Photo</label>
          <div class="photo-upload">
            <div class="photo-preview" id="photoPreview">${photoHtml}</div>
            <div class="photo-controls">
              <label class="btn btn-ghost btn-sm" for="photoInput">Upload Photo</label>
              <input type="file" id="photoInput" accept="image/*" onchange="handlePhotoUpload(event)">
              <div class="hint">JPG/PNG · optional · stored as base64 in browser</div>
            </div>
          </div>
        </div>
        <div class="form-row">
          <div class="form-field"><label>Employee ID <span class="req">*</span></label><input id="f_id" value="${emp?emp.id:''}" ${isEdit?'readonly':''} placeholder="EMP007" required></div>
          <div class="form-field"><label>Full Name <span class="req">*</span></label><input id="f_name" value="${emp?emp.name:''}" placeholder="Enter Name" required></div>
        </div>
        <div class="form-row">
          <div class="form-field"><label>Position <span class="req">*</span></label><input id="f_position" value="${emp?emp.position:''}" placeholder="Enter Position" required></div>
          <div class="form-field"><label>Department <span class="req">*</span></label>
            <select id="f_department" required>${DEPARTMENTS.map(d=>`<option ${emp&&emp.department===d?'selected':''}>${d}</option>`).join('')}</select>
          </div>
        </div>
        <div class="form-row">
          <div class="form-field"><label>Phone Number</label><input id="f_phone" value="${emp?emp.phone:''}" placeholder="+91 ..."></div>
          <div class="form-field"><label>Joining Date <span class="req">*</span></label><input type="date" id="f_joining" value="${emp?emp.joiningDate:''}" required></div>
        </div>
        <div class="form-field"><label>Password (demo only) <span class="req">*</span></label><input id="f_password" value="${emp?emp.password:DEFAULT_EMP_PASS}" placeholder="emp@123" required><div class="hint">Used for employee login. Demo only — not secure.</div></div>
        <div class="modal-foot">
          <button type="button" class="btn btn-ghost" onclick="closeEmpModal()">Cancel</button>
          <button type="submit" class="btn btn-primary">${isEdit?'Save Changes':'Add Employee'}</button>
        </div>
      </form>
    </div>`;
  document.getElementById('empModalOverlay').classList.add('active');
}

function closeEmpModal(){ document.getElementById('empModalOverlay').classList.remove('active'); }

function handlePhotoUpload(e){
  const file = e.target.files[0];
  if(!file) return;
  if(file.size > 1.5*1024*1024){ toast('Image too large','Please use an image under 1.5 MB','warning'); return; }
  const reader = new FileReader();
  reader.onload = ev=>{
    _pendingPhoto = ev.target.result;
    document.getElementById('photoPreview').innerHTML = `<img src="${_pendingPhoto}" alt="">`;
  };
  reader.readAsDataURL(file);
}



async function saveEmployee(e) {
  e.preventDefault();

  const id = document.getElementById('f_id').value.trim();
  const name = document.getElementById('f_name').value.trim();
  const position = document.getElementById('f_position').value.trim();
  const department = document.getElementById('f_department').value;
  const phone = document.getElementById('f_phone').value.trim();
  const joiningDate = document.getElementById('f_joining').value;
  const password = document.getElementById('f_password').value;

  const isEditing = !!_editingEmpId;

  if (!id || !name || (!isEditing && !password)) {
    toast(
      'Missing details',
      isEditing ? 'Employee ID and Name required' : 'Employee ID, Name and Password required',
      'error'
    );
    return false;
  }

  const employeeData = {
    employee_code: isEditing ? _editingEmpId : id,
    name: name,
    email: '',
    phone: phone,
    position: position,
    department: department,
    joining_date: joiningDate,
    password: password
  };

  try {
    const response = await fetch('./employee_api.php', {
      method: isEditing ? 'PUT' : 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      credentials: 'same-origin',
      body: JSON.stringify(employeeData)
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      toast('Error', result.message || 'Employee save failed', 'error');
      return false;
    }

    toast(
      isEditing ? 'Employee Updated' : 'Employee Added',
      name + (isEditing ? ' updated in MySQL successfully' : ' saved to MySQL successfully'),
      'success'
    );

    closeEmpModal();

    await loadEmployeesFromDatabase();

    populateDeptFilter();
    renderEmployees();

  } catch (error) {
    console.error('Employee Save Error:', error);
    toast('Server Error', 'Cannot connect to PHP database API', 'error');
  }

  return false;
}

async function loadEmployeesFromDatabase() {
  try {
    const response = await fetch('./employee_api.php', {
      method: 'GET',
      credentials: 'same-origin'
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.message || 'Employee loading failed');
    }

    const employees = result.data.map(emp => ({
      id: emp.employee_code,
      name: emp.name,
      position: emp.position || '',
      department: emp.department || '',
      phone: emp.phone || '',
      joiningDate: emp.joining_date || '',
      photo: '',
      color: '#2196F3'
    }));

    DB.saveEmployees(employees);

    return employees;

  } catch (error) {
    console.error('Load Employees Error:', error);
    toast('Error', 'Unable to load employees from database', 'error');
    return [];
  }
}


async function deleteEmployee(id) {
  const emp = DB.getEmployees().find(e => e.id === id);

  if (!emp) return;

  if (!confirm(
    `Delete employee "${emp.name}" (${emp.id})?\n\n` +
    `This employee will be removed from the active employee list. ` +
    `Their attendance history will be preserved.`
  )) return;

  try {
    const response = await fetch('./employee_api.php', {
      method: 'DELETE',
      headers: {
        'Content-Type': 'application/json'
      },
      credentials: 'same-origin',
      body: JSON.stringify({
        employee_code: id
      })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      toast('Delete Failed', result.message || 'Employee delete failed', 'error');
      return;
    }

    toast(
      'Employee Deleted',
      emp.name + ' removed successfully',
      'success'
    );

    // Reload active employees from MySQL
    await loadEmployeesFromDatabase();

    populateDeptFilter();
    renderEmployees();

  } catch (error) {
    console.error('Employee Delete Error:', error);
    toast('Server Error', 'Cannot connect to PHP database API', 'error');
  }
}


/* View profile */
function viewEmployee(id){
  const emp = DB.getEmployees().find(e=>e.id===id);
  if(!emp) return;
  const att = DB.getAttendance().filter(a=>a.employeeId===id);
  const present = att.filter(a=>a.status==='Present').length;
  const late = att.filter(a=>a.status==='Late').length;
  const absent = att.filter(a=>a.status==='Absent').length;
  const totalHrs = att.reduce((s,a)=>s+(a.workingHours||0),0);
  const initials = emp.name.split(' ').map(s=>s[0]).join('').slice(0,2).toUpperCase();
  const photo = emp.photo? `<img src="${emp.photo}" alt="">` : initials;
  document.getElementById('profileModal').innerHTML = `
    <div class="modal-head"><h4>Employee Profile</h4><button class="close" onclick="document.getElementById('profileModalOverlay').classList.remove('active')"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button></div>
    <div class="modal-body">
      <div class="profile-grid">
        <div class="profile-photo-wrap">
          <div class="profile-photo-big">${photo}</div>
          <h3 style="font-size:18px;font-weight:700;">${emp.name}</h3>
          <p style="color:var(--text-light);font-size:13px;">${emp.position}</p>
          <div style="margin-top:8px;">${badgeHtml(present>=late?'Present':'Late')}</div>
        </div>
        <div>
          <div class="profile-detail-row"><span class="pdr-label">Employee ID</span><span class="pdr-value cell-id">${emp.id}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Department</span><span class="pdr-value">${emp.department}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Position</span><span class="pdr-value">${emp.position}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Phone</span><span class="pdr-value">${emp.phone||'—'}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Joining Date</span><span class="pdr-value">${prettyDate(emp.joiningDate)}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Total Records</span><span class="pdr-value">${att.length}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Present Days</span><span class="pdr-value" style="color:var(--accent)">${present}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Late Days</span><span class="pdr-value" style="color:var(--warning)">${late}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Absent Days</span><span class="pdr-value" style="color:var(--danger)">${absent}</span></div>
          <div class="profile-detail-row"><span class="pdr-label">Total Hours</span><span class="pdr-value">${round2(totalHrs)} h</span></div>
        </div>
      </div>
    </div>`;
  document.getElementById('profileModalOverlay').classList.add('active');
}

/* ---------- Attendance table ---------- */

async function renderAttendanceTable() {
  const body = document.getElementById('attTableBody');

  body.innerHTML = `
    <tr>
      <td colspan="8">Loading attendance records...</td>
    </tr>
  `;

  try {
    const response = await fetch('./admin_attendance_api.php', {
      method: 'GET',
      credentials: 'same-origin'
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.message || 'Attendance loading failed');
    }

    const q = (document.getElementById('attSearch').value || '').toLowerCase();
    const df = document.getElementById('attDateFilter').value;
    const sf = document.getElementById('attStatusFilter').value;

    const recs = result.data.map(r => {
      const hours = Number(r.total_hours || 0);

      return {
        id: r.id,
        employeeId: r.employee_code,
        date: r.attendance_date,
        checkIn: r.check_in,
        checkOut: r.check_out,
        workingHours: hours,
        overtimeHours: Math.max(0, round2(hours - STANDARD_HOURS)),
        status: r.status,
        employee: {
          id: r.employee_code,
          name: r.name,
          position: r.position || '',
          department: r.department || '',
          phone: r.phone || '',
          photo: '',
          color: '#2196F3'
        }
      };
    });

    let filtered = recs;

    if (df) {
      filtered = filtered.filter(r => r.date === df);
    }

    if (sf) {
      filtered = filtered.filter(r => r.status === sf);
    }

    if (q) {
      filtered = filtered.filter(r =>
        r.employee.name.toLowerCase().includes(q) ||
        r.employee.id.toLowerCase().includes(q)
      );
    }

    if (filtered.length === 0) {
      body.innerHTML = `
        <tr>
          <td colspan="8">
            <div class="empty">
              <h4>No records</h4>
              <p>No attendance records match your filters.</p>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    body.innerHTML = filtered.map(r => {
      const e = r.employee;

      return `<tr>
        <td>
          <div class="avatar-cell">
            ${avatarHtml(e)}
            <div>
              <div class="cell-strong">${e.name}</div>
              <div class="cell-muted">${e.position}</div>
            </div>
          </div>
        </td>
        <td class="cell-id">${e.id}</td>
        <td>${prettyDate(r.date)}</td>
        <td>${r.checkIn || '<span class="cell-muted">—</span>'}</td>
        <td>${r.checkOut || '<span class="cell-muted">—</span>'}</td>
        <td>${r.checkOut ? (r.workingHours + 'h') : '<span class="cell-muted">—</span>'}</td>
        <td>${r.checkOut && r.overtimeHours > 0 ? (r.overtimeHours + 'h') : '<span class="cell-muted">—</span>'}</td>
        <td>${badgeHtml(r.status)}</td>
      </tr>`;
    }).join('');

  } catch (error) {
    console.error('Attendance Table Error:', error);

    body.innerHTML = `
      <tr>
        <td colspan="8">
          <div class="empty">
            <h4>Unable to load records</h4>
            <p>${error.message}</p>
          </div>
        </td>
      </tr>
    `;
  }
}

function markAbsentToday(){
  const emps=DB.getEmployees();
  const att=DB.getAttendance();
  const today=todayDateStr();
  let added=0;
  emps.forEach(e=>{
    const exists = att.some(a=>a.employeeId===e.id && a.date===today);
    if(!exists){
      att.push({ id:recId(), employeeId:e.id, date:today, checkIn:null, checkOut:null, workingHours:0, overtimeHours:0, status:'Absent' });
      added++;
    }
  });
  DB.saveAttendance(att);
  toast('Absent marked', added+' employees without check-in marked absent','warning');
  renderAttendanceTable();
}

/* ---------- Employee app ---------- */
let _currentEmp = null;
function initEmployeeApp(){
  const s = DB.getSession();
  _currentEmp = DB.getEmployees().find(e=>e.id===s.id);
  if(!_currentEmp){ logout(); return; }
  renderEmpHero();
  renderEmpUserMini();
  updateEmpAttendanceUI();
  renderEmpHistory();
}

function renderEmpHero(){
  const e=_currentEmp;
  document.getElementById('empHero').innerHTML = `
    <div class="emp-info">
      ${avatarHtml(e,'avatar',64)}
      <div><h3>${e.name}</h3><p>${e.position} · ${e.department}</p><span class="id-pill">${e.id}</span></div>
    </div>
    <div class="clock-hero"><div class="big-time">--:--:--</div><div class="big-date">—</div></div>`;
}

function renderEmpUserMini(){
  const e=_currentEmp;
  document.getElementById('empUserMini').innerHTML = `${avatarHtml(e,'avatar',36)}<div class="uinfo"><strong>${e.name}</strong><span>${e.id}</span></div>`;
}

function getTodayRecord(){
  const today=todayDateStr();
  return DB.getAttendance().find(a=>a.employeeId===_currentEmp.id && a.date===today);
}

function updateEmpAttendanceUI(){
  const rec=getTodayRecord();
  const ciBtn=document.getElementById('checkinBtn');
  const coBtn=document.getElementById('checkoutBtn');
  const ciT=document.getElementById('checkinTime');
  const coT=document.getElementById('checkoutTime');
  if(rec && rec.checkIn){
    ciBtn.disabled=true;
    ciT.textContent='Checked in at '+rec.checkIn;
    if(rec.checkOut){
      coBtn.disabled=true;
      coT.textContent='Checked out at '+rec.checkOut;
    } else {
      coBtn.disabled=false;
      coT.textContent='Ready to check out';
    }
  } else {
    ciBtn.disabled=false;
    ciT.textContent='Not checked in';
    coBtn.disabled=true;
    coT.textContent='Check in first';
  }
  renderTodaySummary();
  renderEmpHistory();
}

function renderTodaySummary(){
  const rec=getTodayRecord();
  const tiles = [
    {label:'Entry Time', value:(rec&&rec.checkIn)?rec.checkIn:'—'},
    {label:'Exit Time', value:(rec&&rec.checkOut)?rec.checkOut:'—'},
    {label:'Working Hours', value:(rec&&rec.checkOut)?(rec.workingHours+'h'):'—', cls:rec&&rec.checkOut?'accent':''},
    {label:'Overtime', value:(rec&&rec.checkOut&&rec.overtimeHours>0)?(rec.overtimeHours+'h'):'0h', cls:rec&&rec.checkOut&&rec.overtimeHours>0?'warning':''},
    {label:'Status', value:rec?rec.status:'Not marked'},
    {label:'Standard Hours', value:STANDARD_HOURS+'h'}
  ];
  document.getElementById('todayStatus').innerHTML = tiles.map(t=>`<div class="status-tile"><div class="st-label">${t.label}</div><div class="st-value ${t.cls||''}">${t.value}</div></div>`).join('');
}

async function syncTodayAttendance() {
  const response = await fetch('./attendance_api.php', {
    method: 'GET',
    credentials: 'same-origin'
  });

  const result = await response.json();

  if (!response.ok || !result.success) {
    throw new Error(result.message || 'Attendance loading failed');
  }

  const att = DB.getAttendance();
  const today = todayDateStr();

  // Remove old cached record for this employee and today
  const filtered = att.filter(a =>
    !(a.employeeId === _currentEmp.id && a.date === today)
  );

  if (result.data) {
    const record = result.data;

    filtered.push({
      id: record.id,
      employeeId: _currentEmp.id,
      date: record.attendance_date,
      checkIn: record.check_in,
      checkOut: record.check_out,
      workingHours: Number(record.total_hours || 0),
      overtimeHours: 0,
      status: record.status
    });
  }

  DB.saveAttendance(filtered);
}



async function empCheckIn() {
  try {
    const response = await fetch('./attendance_api.php', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'same-origin',
      body: JSON.stringify({ action: 'check_in' })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      toast('Check-in failed', result.message, 'error');
      return;
    }

    await syncTodayAttendance();
    updateEmpAttendanceUI();

    const timeStr = formatTime(new Date());
    const status = result.data.status;

    toast(
      'Check-in successful',
      'Time: ' + timeStr + ' · Status: ' + status,
      status === 'Late' ? 'warning' : 'success'
    );

  } catch (error) {
    console.error(error);
    toast('Server Error', 'Unable to save check-in', 'error');
  }
}



async function empCheckOut() {
  try {
    const response = await fetch('./attendance_api.php', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'same-origin',
      body: JSON.stringify({ action: 'check_out' })
    });

    const result = await response.json();

    if (!response.ok || !result.success) {
      toast('Check-out failed', result.message, 'error');
      return;
    }

    await syncTodayAttendance();
    updateEmpAttendanceUI();

    const worked = Number(result.data.total_hours);

    toast(
      'Check-out successful',
      'Worked ' + worked.toFixed(2) + ' hours',
      'success'
    );

  } catch (error) {
    console.error(error);
    toast('Server Error', 'Unable to save check-out', 'error');
  }
}


function renderEmpHistory(){
  const att=DB.getAttendance().filter(a=>a.employeeId===_currentEmp.id).sort((a,b)=>a.date<b.date?1:-1);
  const body=document.getElementById('empHistoryBody');
  if(att.length===0){ body.innerHTML='<tr><td colspan="6"><div class="empty"><h4>No history yet</h4><p>Your attendance records will appear here.</p></div></td></tr>'; return; }
  body.innerHTML=att.map(r=>`<tr>
    <td>${prettyDate(r.date)}</td>
    <td>${r.checkIn||'<span class="cell-muted">—</span>'}</td>
    <td>${r.checkOut||'<span class="cell-muted">—</span>'}</td>
    <td>${r.checkOut?(r.workingHours+'h'):'<span class="cell-muted">—</span>'}</td>
    <td>${r.checkOut&&r.overtimeHours>0?(r.overtimeHours+'h'):'<span class="cell-muted">—</span>'}</td>
    <td>${badgeHtml(r.status)}</td>
  </tr>`).join('');
}

/* ---------- Exports ---------- */
function exportAttendanceExcel(){
  const att=DB.getAttendance();
  const emps=DB.getEmployees(); const byId={}; emps.forEach(e=>byId[e.id]=e);
  const rows=att.map(a=>{ const e=byId[a.employeeId]||{}; return {
    'Employee ID':a.employeeId,'Employee Name':e.name||'','Department':e.department||'',
    'Date':a.date,'Check-In':a.checkIn||'','Check-Out':a.checkOut||'',
    'Working Hours':a.workingHours||0,'Overtime Hours':a.overtimeHours||0,'Status':a.status
  };}).sort((a,b)=>a.Date<b.Date?1:-1);
  if(rows.length===0){ toast('No data','No attendance records to export','warning'); return; }
  const ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=[{wch:12},{wch:22},{wch:14},{wch:12},{wch:12},{wch:12},{wch:14},{wch:14},{wch:12}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Attendance');
  XLSX.writeFile(wb,'YUVAPACK_Attendance_'+todayDateStr()+'.xlsx');
  toast('Excel exported','Attendance records downloaded','success');
}

function exportEmployeesExcel(){
  const emps=DB.getEmployees();
  if(emps.length===0){ toast('No data','No employees to export','warning'); return; }
  const rows=emps.map(e=>({ 'Employee ID':e.id,'Name':e.name,'Position':e.position,'Department':e.department,'Phone':e.phone||'','Joining Date':e.joiningDate,'Password':e.password }));
  const ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=[{wch:12},{wch:22},{wch:22},{wch:14},{wch:18},{wch:14},{wch:12}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Employees');
  XLSX.writeFile(wb,'YUVAPACK_Employees_'+todayDateStr()+'.xlsx');
  toast('Excel exported','Employee directory downloaded','success');
}

function exportAttendancePDF(){
  const att=DB.getAttendance();
  const emps=DB.getEmployees(); const byId={}; emps.forEach(e=>byId[e.id]=e);
  if(att.length===0){ toast('No data','No attendance records to export','warning'); return; }
  const { jsPDF } = window.jspdf;
  const doc=new jsPDF({orientation:'landscape',unit:'pt',format:'a4'});
  doc.setFontSize(16); doc.setTextColor(33,150,243);
  doc.text('YUVAPACK — Attendance Report',40,40);
  doc.setFontSize(10); doc.setTextColor(96,125,139);
  doc.text('Generated: '+new Date().toLocaleString(),40,58);
  doc.text('Total records: '+att.length,40,72);
  const rows=att.slice().sort((a,b)=>a.date<b.date?1:-1).map(a=>{
    const e=byId[a.employeeId]||{};
    return [a.employeeId, e.name||'', a.date, a.checkIn||'-', a.checkOut||'-', (a.workingHours||0)+'h', (a.overtimeHours||0)+'h', a.status];
  });
  doc.autoTable({
    head:[['ID','Name','Date','Check-In','Check-Out','Hours','Overtime','Status']],
    body:rows, startY:88, margin:{left:40,right:40},
    styles:{fontSize:8.5,cellPadding:5},
    headStyles:{fillColor:[33,150,243],textColor:255},
    alternateRowStyles:{fillColor:[238,243,248]},
    columnStyles:{0:{cellWidth:55},1:{cellWidth:120}}
  });
  doc.save('YUVAPACK_Attendance_'+todayDateStr()+'.pdf');
  toast('PDF exported','Attendance report downloaded','success');
}

function exportSummaryPDF(){
  const emps=DB.getEmployees();
  const att=DB.getAttendance().filter(a=>a.date===todayDateStr());
  let present=0,late=0,absent=0,halfday=0,leave=0;
  const ids=new Set();
  att.forEach(a=>{ ids.add(a.employeeId); if(a.status==='Present')present++; else if(a.status==='Late')late++; else if(a.status==='Half Day')halfday++; else if(a.status==='On Leave')leave++; else if(a.status==='Absent')absent++; });
  absent += emps.length - ids.size;
  const totalHours=att.reduce((s,a)=>s+(a.workingHours||0),0);
  const { jsPDF }=window.jspdf;
  const doc=new jsPDF({unit:'pt',format:'a4'});
  // header band
  doc.setFillColor(33,150,243); doc.rect(0,0,595,70,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(20); doc.text('YUVAPACK',40,38);
  doc.setFontSize(11); doc.text('Employee Attendance Management System',40,54);
  doc.setTextColor(96,125,139); doc.setFontSize(10);
  doc.text('Dashboard Summary — '+new Date().toLocaleDateString('en-US',{weekday:'long',year:'numeric',month:'long',day:'numeric'}),40,96);
  doc.setFontSize(9); doc.text('Generated: '+new Date().toLocaleString(),40,110);
  const stats=[
    ['Total Employees',emps.length],['Present Today',present],['Late Today',late],
    ['Absent Today',absent],['On Leave',leave],['Half Day',halfday],['Total Working Hours',round2(totalHours)+'h']
  ];
  doc.autoTable({
    head:[['Metric','Value']],
    body:stats, startY:128, margin:{left:40,right:40},
    styles:{fontSize:11,cellPadding:8},
    headStyles:{fillColor:[33,150,243],textColor:255},
    alternateRowStyles:{fillColor:[238,243,248]},
    columnStyles:{0:{cellWidth:300},1:{cellWidth:200,halign:'right'}}
  });
  let y=doc.lastAutoTable.finalY+24;
  doc.setFontSize(10); doc.setTextColor(38,50,56);
  doc.text('Office Timing Policy',40,y);
  doc.setFontSize(9); doc.setTextColor(96,125,139);
  y+=16;
  doc.text('• Office Start: 09:00 AM',40,y); y+=13;
  doc.text('• Grace Period: 15 minutes (until 09:15 AM)',40,y); y+=13;
  doc.text('• Late After: 09:15 AM  →  status: Late',40,y); y+=13;
  doc.text('• After 12:00 PM  →  status: Half Day',40,y); y+=13;
  doc.text('• No check-in  →  status: Absent',40,y); y+=13;
  doc.text('• Standard working hours: 8h  (overtime = worked − 8)',40,y);
  doc.save('YUVAPACK_Summary_'+todayDateStr()+'.pdf');
  toast('PDF exported','Dashboard summary downloaded','success');
}

/* ---------- Helpers: icons & avatar ---------- */
function avatarHtml(emp, cls='avatar', size=36){
  if(emp.photo) return `<div class="${cls}" style="width:${size}px;height:${size}px;background:${emp.color||'#2196F3'}"><img src="${emp.photo}" alt=""></div>`;
  const initials = emp.name.split(' ').map(s=>s[0]).join('').slice(0,2).toUpperCase();
  return `<div class="${cls}" style="width:${size}px;height:${size}px;background:${emp.color||'#2196F3'};font-size:${Math.round(size*0.4)}px">${initials}</div>`;
}
function badgeHtml(status){
  const map={'Present':'badge-present','Late':'badge-late','Absent':'badge-absent','Half Day':'badge-halfday','On Leave':'badge-leave'};
  return `<span class="badge ${map[status]||'badge-present'}">${status}</span>`;
}
function empIcon(){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>';}
function checkIcon(){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>';}
function clockIcon(){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>';}
function xIcon(){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>';}
function planeIcon(){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17.8 19.2L16 11l3.5-3.5C21 6 21.5 4 21 3c-1-.5-3 0-4.5 1.5L13 8 4.8 6.2c-.5-.1-.9.1-1.1.5l-.3.5c-.2.5-.1 1 .3 1.3L9 12l-2 3H4l-1 1 3 2 2 3 1-1v-3l3-2 3.5 5.3c.3.4.8.5 1.3.3l.5-.2c.4-.3.6-.7.5-1.2z"/></svg>';}
function hourIcon(){return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 9 15"/></svg>';}

/* ---------- Sidebar mobile ---------- */
function openSidebar(){ document.getElementById('sidebar').classList.add('open'); document.getElementById('backdrop').classList.add('active'); }
function closeSidebar(){ document.getElementById('sidebar').classList.remove('open'); document.getElementById('backdrop').classList.remove('active'); }

/* ---------- Init ---------- */
(function init(){
  seedData();
  startClock();
  // close modals on overlay click
  document.getElementById('empModalOverlay').addEventListener('click',e=>{ if(e.target.id==='empModalOverlay') closeEmpModal(); });
  document.getElementById('profileModalOverlay').addEventListener('click',e=>{ if(e.target.id==='profileModalOverlay') e.target.classList.remove('active'); });
  // restore session if present
  const s=DB.getSession();
  if(s){ showApp(s.role); }
  // (date filter intentionally left empty so all records show by default)
})();
</script>
</body>
</html>
