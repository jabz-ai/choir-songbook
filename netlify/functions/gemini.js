<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Choir Song Book — Brethren Assembly Abu Dhabi</title>

<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,500;0,600;0,700;1,400&family=Crimson+Text:ital,wght@0,400;0,600;1,400&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600;9..40,700&display=swap" rel="stylesheet">

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-app-compat.js"></script>
<script src="https://www.gstatic.com/firebasejs/10.7.0/firebase-firestore-compat.js"></script>

<style>
:root {
  --bg:#070f1c; --surface:#0c1828; --card:#0f1e30; --card-2:#0d1b2d;
  --border:rgba(196,151,42,0.18); --border-l:rgba(196,151,42,0.38); --border-strong:rgba(196,151,42,0.6);
  --gold:#C4972A; --gold-lt:#E8C36A; --gold-dk:#9a7520;
  --cream:#F2EFE6; --muted:#8A97B0; --muted-2:#4e5f78;
  --danger:#c0404a; --success:#2e9e65;
  --r-sm:8px; --r-md:12px; --r-lg:16px; --r-xl:22px;
  --glow:0 0 24px rgba(196,151,42,0.08), 0 8px 32px rgba(0,0,0,0.5);
  --glow-hover:0 0 32px rgba(196,151,42,0.15), 0 12px 40px rgba(0,0,0,0.55);
  --glow-gold:0 0 20px rgba(196,151,42,0.25), 0 4px 16px rgba(0,0,0,0.4);
  --shadow:0 8px 32px rgba(0,0,0,0.5); --shadow-sm:0 4px 16px rgba(0,0,0,0.35);
  --serif:'Playfair Display',Georgia,serif;
  --body:'DM Sans',system-ui,sans-serif;
  --lyrics:'Crimson Text',Georgia,serif;
  --t:all 0.25s cubic-bezier(.4,0,.2,1);
}
*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{font-family:var(--body);background:var(--bg);color:var(--cream);min-height:100vh;line-height:1.6;-webkit-font-smoothing:antialiased;background-image:radial-gradient(ellipse at 20% 0%,rgba(196,151,42,0.04) 0%,transparent 60%),radial-gradient(ellipse at 80% 100%,rgba(196,151,42,0.03) 0%,transparent 50%)}

/* ── HEADER ── */
.app-header{position:sticky;top:0;z-index:50;background:rgba(7,15,28,0.95);backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:18px;padding:14px 32px;box-shadow:0 1px 0 rgba(196,151,42,0.1),0 4px 24px rgba(0,0,0,0.4)}
.brand-block{display:flex;align-items:center;gap:14px}
.brand-icon{width:38px;height:38px;background:linear-gradient(145deg,var(--gold-lt),var(--gold),var(--gold-dk));border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:17px;box-shadow:0 0 16px rgba(196,151,42,0.35),0 2px 8px rgba(0,0,0,0.4)}
.brand-text{display:flex;flex-direction:column;line-height:1.2}
.brand-title{font-family:var(--serif);font-size:16px;font-weight:600;color:var(--gold-lt);letter-spacing:-.01em}
.brand-sub{font-size:9px;color:var(--muted);letter-spacing:.13em;text-transform:uppercase;font-weight:500;margin-top:2px}

.nav-pill{background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:var(--r-lg);padding:5px;display:inline-flex;gap:3px;box-shadow:0 0 0 1px rgba(196,151,42,0.05) inset}
.nav-item{padding:8px 20px;color:var(--muted);border-radius:10px;font-size:13px;font-weight:500;cursor:pointer;transition:var(--t);background:none;border:none;font-family:var(--body);letter-spacing:.01em}
.nav-item:hover{color:var(--cream);background:rgba(255,255,255,0.05)}
.nav-item.on{background:linear-gradient(135deg,var(--gold),var(--gold-dk));color:#07111e;font-weight:700;box-shadow:0 0 16px rgba(196,151,42,0.4),0 2px 8px rgba(0,0,0,0.3)}

.header-right{display:flex;align-items:center;gap:10px}
.quick-find{display:flex;align-items:center;gap:8px;background:rgba(255,255,255,0.03);border:1px solid var(--border);padding:8px 14px;border-radius:10px;font-size:12px;color:var(--muted);width:210px;transition:var(--t)}
.quick-find:focus-within{border-color:var(--gold);background:rgba(255,255,255,0.05);box-shadow:0 0 12px rgba(196,151,42,0.12)}
.quick-find input{background:none;border:none;outline:none;color:var(--cream);font-family:var(--body);font-size:12px;flex:1;width:100%}
.quick-find input::placeholder{color:var(--muted)}
.status-pill{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--muted);padding:7px 12px;background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:100px}
.status-dot{width:6px;height:6px;border-radius:50%;background:var(--muted);transition:var(--t)}
.status-dot.live{background:var(--success);box-shadow:0 0 8px rgba(46,158,101,0.7)}

/* ── MAIN ── */
.container{max-width:1300px;margin:0 auto;padding:36px 32px 70px}
.view-section{display:none;animation:fadeUp 0.35s ease}
.view-section.active{display:block}
@keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}

/* ── PAGE TITLE BLOCK ── */
.page-tag{font-size:11px;color:var(--gold);letter-spacing:.12em;text-transform:uppercase;font-weight:700;margin-bottom:10px;display:flex;align-items:center;gap:8px}
.page-tag::before{content:'';width:20px;height:2px;background:linear-gradient(90deg,var(--gold),transparent);border-radius:2px}
.page-h1{font-family:var(--serif);font-size:48px;font-weight:700;color:var(--cream);line-height:1;letter-spacing:-.02em;margin-bottom:26px;text-shadow:0 2px 20px rgba(196,151,42,0.1)}

/* ── SEARCH BLOCK ── */
.search-block{background:linear-gradient(160deg,var(--card),var(--card-2));border:1px solid var(--border);border-radius:var(--r-lg);padding:20px 24px;margin-bottom:28px;box-shadow:var(--glow)}
.search-block-label{font-size:11px;color:var(--muted);letter-spacing:.09em;text-transform:uppercase;font-weight:600;margin-bottom:12px}
.search-row{display:flex;gap:10px;align-items:stretch}
.search-input-wrap{flex:1;position:relative}
.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--muted);font-size:14px;pointer-events:none}
.search-input{width:100%;padding:14px 16px 14px 42px;background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:var(--r-md);font-family:var(--body);font-size:14px;color:var(--cream);transition:var(--t)}
.search-input::placeholder{color:var(--muted)}
.search-input:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(196,151,42,0.12),0 0 16px rgba(196,151,42,0.08);background:rgba(255,255,255,0.05)}

/* ── LANGUAGE FILTER ── */
.lang-filter{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:30px;padding-top:8px;align-items:center}
.lang-chip{padding:7px 18px;border-radius:100px;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;font-size:12px;font-weight:500;transition:var(--t);font-family:var(--body)}
.lang-chip:hover{border-color:var(--gold);color:var(--gold-lt);background:rgba(196,151,42,0.06)}
.lang-chip.active{background:linear-gradient(135deg,rgba(196,151,42,0.2),rgba(196,151,42,0.1));border-color:var(--gold);color:var(--gold-lt);font-weight:600;box-shadow:0 0 12px rgba(196,151,42,0.15)}
.reset-all{margin-left:auto;background:none;border:none;color:var(--muted);font-size:11px;letter-spacing:.08em;text-transform:uppercase;cursor:pointer;font-family:var(--body);font-weight:500;display:inline-flex;align-items:center;gap:5px;transition:var(--t)}
.reset-all:hover{color:var(--cream)}

/* ── LANGUAGE GROUP ── */
.lang-group{margin-bottom:42px}
.lang-group-head{display:flex;align-items:baseline;gap:14px;margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid var(--border)}
.lang-group-title{font-family:var(--serif);font-size:22px;font-weight:600;color:var(--cream)}
.lang-group-count{font-size:11px;color:var(--gold);background:rgba(196,151,42,0.1);border:1px solid rgba(196,151,42,0.22);padding:3px 11px;border-radius:100px;letter-spacing:.04em;font-weight:600}

/* ── SONG CARDS ── */
.songs-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:16px}
.song-card{position:relative;background:linear-gradient(160deg,var(--card),var(--card-2));border:1px solid var(--border);border-radius:var(--r-lg);padding:20px 20px 17px;cursor:pointer;transition:var(--t);overflow:hidden;box-shadow:var(--glow)}
.song-card::before{content:'';position:absolute;top:0;left:0;right:0;height:1px;background:linear-gradient(90deg,transparent,rgba(196,151,42,0.2),transparent);opacity:0;transition:var(--t)}
.song-card:hover{border-color:var(--border-l);transform:translateY(-4px);box-shadow:var(--glow-hover)}
.song-card:hover::before{opacity:1}
.song-card-tags{display:flex;gap:5px;margin-bottom:12px;flex-wrap:wrap}
.song-tag{font-size:9px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;padding:3px 9px;border-radius:5px;color:var(--gold-lt);background:rgba(196,151,42,0.1);border:1px solid rgba(196,151,42,0.22)}
.song-tag.serial{color:var(--gold);background:rgba(196,151,42,0.06)}
.song-tag.english{color:#5cd48a;background:rgba(46,158,101,0.12);border-color:rgba(46,158,101,0.25)}
.song-tag.video{color:#f27680;background:rgba(192,64,74,0.12);border-color:rgba(192,64,74,0.25)}
.song-card-title{font-family:var(--serif);font-size:18px;font-weight:700;color:var(--cream);line-height:1.25;margin-bottom:4px;letter-spacing:-.005em}
.song-card-artist{font-size:11px;color:var(--muted);letter-spacing:.05em;text-transform:uppercase;font-weight:500;margin-bottom:14px}
.song-card-foot{display:flex;align-items:center;justify-content:space-between;padding-top:12px;border-top:1px solid var(--border)}
.song-card-meta{font-size:11px;color:var(--muted);display:flex;align-items:center;gap:6px}
.song-card-arrow{color:var(--gold);font-size:14px;transition:var(--t)}
.song-card:hover .song-card-arrow{transform:translateX(4px)}

/* ── EMPTY/LOADING STATES ── */
.empty-state{text-align:center;padding:80px 20px;color:var(--muted)}
.empty-state h3{font-family:var(--serif);font-size:22px;color:var(--cream);margin-bottom:8px}
.empty-state p{font-size:14px}
.empty-icon{font-size:42px;opacity:0.4;margin-bottom:18px}
.loading-state{text-align:center;padding:60px}
.spinner{width:32px;height:32px;border-radius:50%;border:3px solid rgba(196,151,42,0.2);border-top-color:var(--gold);animation:spin 0.8s linear infinite;margin:0 auto 16px;box-shadow:0 0 12px rgba(196,151,42,0.1)}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── MODAL ── */
.modal{display:none;position:fixed;inset:0;background:rgba(4,9,18,0.92);backdrop-filter:blur(14px);z-index:1000;justify-content:center;align-items:center;padding:20px}
.modal.active{display:flex;animation:fadeIn 0.22s ease}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
.modal-content{background:linear-gradient(180deg,#101f32,var(--card) 140px);border:1px solid var(--border-l);border-top:3px solid var(--gold);border-radius:var(--r-xl);padding:34px 40px;max-width:1000px;width:100%;max-height:92vh;overflow-y:auto;position:relative;box-shadow:0 0 60px rgba(196,151,42,0.1),0 32px 80px rgba(0,0,0,0.8)}
.close-btn{position:absolute;top:18px;right:18px;width:36px;height:36px;border-radius:50%;border:1px solid var(--border);background:rgba(255,255,255,0.05);color:var(--muted);cursor:pointer;font-size:1.1em;display:flex;align-items:center;justify-content:center;transition:var(--t)}
.close-btn:hover{background:rgba(255,255,255,0.1);color:var(--cream);transform:rotate(90deg);border-color:var(--border-l)}

.modal-back-row{display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;flex-wrap:wrap;gap:10px}
.modal-back{background:none;border:none;color:var(--muted);font-size:13px;cursor:pointer;font-family:var(--body);display:inline-flex;align-items:center;gap:6px;padding:6px 0;transition:var(--t)}
.modal-back:hover{color:var(--gold-lt)}
.modal-actions{display:flex;gap:8px;flex-wrap:wrap}

.modal-tags-row{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}
.modal-song-title{font-family:var(--serif);font-size:36px;font-weight:700;color:var(--cream);margin-bottom:6px;line-height:1.1;letter-spacing:-.015em;text-shadow:0 2px 20px rgba(196,151,42,0.1)}
.modal-song-artist{font-size:12px;color:var(--muted);letter-spacing:.08em;text-transform:uppercase;font-weight:500;margin-bottom:24px;padding-bottom:14px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px}
.modal-song-artist::before{content:'';width:24px;height:1px;background:linear-gradient(90deg,var(--gold),transparent)}

/* Lyrics display */
.lyrics-display{font-family:var(--lyrics);font-size:1.18em;line-height:2.1;color:#e8e5dc;background:rgba(0,0,0,0.2);border:1px solid var(--border);border-radius:var(--r-md);padding:26px 30px;white-space:pre-wrap;margin-top:14px;box-shadow:inset 0 1px 0 rgba(196,151,42,0.06)}
.lyrics-table-head{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:14px}
.lyrics-table-label{font-size:12px;color:var(--muted);font-weight:500;letter-spacing:.04em}
.toggle-view{display:flex;gap:4px;background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:var(--r-md);padding:3px}
.toggle-btn{padding:6px 13px;border-radius:var(--r-sm);border:none;background:transparent;color:var(--muted);font-size:11px;font-weight:500;cursor:pointer;transition:var(--t);font-family:var(--body);letter-spacing:.02em}
.toggle-btn:hover{color:var(--cream)}
.toggle-btn.active{background:rgba(196,151,42,0.18);color:var(--gold-lt);box-shadow:0 0 8px rgba(196,151,42,0.1)}

.lyrics-cols-grid{display:grid;grid-template-columns:1fr 1fr;gap:0;background:rgba(0,0,0,0.2);border:1px solid var(--border);border-radius:var(--r-md);overflow:hidden}
.lyrics-col{padding:24px 26px}
.lyrics-col-left{border-right:1px solid var(--border)}
.lyrics-col-label{font-size:10px;color:var(--gold);letter-spacing:.12em;text-transform:uppercase;font-weight:600;margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--border)}
.lyrics-col-content{font-family:var(--lyrics);font-size:1.05em;line-height:2;color:#e0ddd4;white-space:pre-wrap}
.lyrics-col-right .lyrics-col-content{color:var(--muted)}

.section-header-row td{background:rgba(196,151,42,0.15) !important;color:var(--gold) !important;text-align:center;font-weight:700;letter-spacing:.05em;font-size:.9em;text-transform:uppercase}

/* Reference material */
.ref-material{margin-top:28px;padding-top:24px;border-top:1px solid var(--border)}
.ref-material-label{font-size:11px;color:var(--muted);letter-spacing:.1em;text-transform:uppercase;font-weight:600;margin-bottom:14px;display:flex;align-items:center;gap:10px}
.ref-material-label::before{content:'';width:3px;height:14px;background:linear-gradient(180deg,var(--gold),var(--gold-dk));border-radius:2px}
.ref-card{background:rgba(0,0,0,0.25);border:1px solid var(--border);border-radius:var(--r-md);overflow:hidden;max-width:520px;box-shadow:var(--glow)}
.ref-card-video{position:relative;padding-bottom:56.25%;height:0;overflow:hidden;background:#000}
.ref-card-video iframe{position:absolute;top:0;left:0;width:100%;height:100%;border:none}
.ref-card-body{padding:14px 18px;display:flex;align-items:center;gap:10px;font-size:12px;color:var(--muted)}
.ref-card-body::before{content:'▶';color:#f27680;font-size:11px}

/* ── BUTTONS ── */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:7px;padding:11px 22px;border:none;border-radius:var(--r-md);cursor:pointer;font-family:var(--body);font-size:13px;font-weight:600;letter-spacing:.01em;transition:var(--t)}
.btn-primary{background:linear-gradient(135deg,var(--gold),var(--gold-dk));color:#07111e;box-shadow:0 0 20px rgba(196,151,42,0.3),0 3px 12px rgba(0,0,0,0.3)}
.btn-primary:hover{filter:brightness(1.1);transform:translateY(-1px);box-shadow:0 0 28px rgba(196,151,42,0.4),0 6px 16px rgba(0,0,0,0.35)}
.btn-primary:disabled{filter:none;transform:none;opacity:0.6;cursor:not-allowed;box-shadow:none}
.btn-secondary{background:rgba(255,255,255,0.05);color:var(--cream);border:1px solid var(--border-l)}
.btn-secondary:hover{background:rgba(255,255,255,0.09);border-color:var(--gold);box-shadow:0 0 12px rgba(196,151,42,0.1)}
.btn-danger{background:rgba(192,64,74,0.15);color:#f27680;border:1px solid rgba(192,64,74,0.3)}
.btn-danger:hover{background:rgba(192,64,74,0.25);box-shadow:0 0 12px rgba(192,64,74,0.15)}
.btn-sm{padding:7px 14px;font-size:11px}
.btn:disabled{opacity:0.5;cursor:not-allowed}
.btn-icon{display:inline-flex;align-items:center;gap:6px}

.btn-sparkle{background:linear-gradient(135deg,rgba(196,151,42,0.15),rgba(196,151,42,0.05));border:1px solid var(--border-l);color:var(--gold-lt);font-size:11px;padding:6px 13px;border-radius:var(--r-sm);cursor:pointer;font-family:var(--body);font-weight:600;transition:var(--t);display:inline-flex;align-items:center;gap:5px}
.btn-sparkle:hover{background:rgba(196,151,42,0.25);color:var(--cream);border-color:var(--gold);box-shadow:0 0 12px rgba(196,151,42,0.15)}
.btn-sparkle:disabled{opacity:0.5;cursor:not-allowed}

/* ── ADMIN ── */
.admin-page-head{display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:28px;flex-wrap:wrap;gap:14px}
.admin-actions-top{display:flex;gap:8px;flex-wrap:wrap}

.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:14px;margin-bottom:28px}
.stat-card{background:linear-gradient(160deg,var(--card),var(--card-2));border:1px solid var(--border);border-radius:var(--r-lg);padding:20px 22px;position:relative;overflow:hidden;transition:var(--t);text-align:center;box-shadow:var(--glow)}
.stat-card:hover{border-color:var(--border-l);transform:translateY(-3px);box-shadow:var(--glow-hover)}
.stat-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--gold),var(--gold-lt),transparent)}
.stat-label{font-size:10px;color:var(--muted);font-weight:600;letter-spacing:.09em;text-transform:uppercase;margin-bottom:10px}
.stat-number{font-family:var(--serif);font-size:36px;font-weight:700;color:var(--gold);line-height:1;letter-spacing:-.01em;text-shadow:0 0 20px rgba(196,151,42,0.3)}
.stat-number.positive{color:#5cd48a;text-shadow:0 0 20px rgba(46,158,101,0.3)}

.admin-layout{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,1fr);gap:20px;align-items:start}
.admin-side-stack{display:flex;flex-direction:column;gap:16px}

.panel-card{background:linear-gradient(160deg,var(--card),var(--card-2));border:1px solid var(--border);border-radius:var(--r-lg);overflow:hidden;box-shadow:var(--glow);transition:var(--t)}
.panel-card:hover{box-shadow:var(--glow-hover)}
.panel-card-head{padding:16px 22px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;background:rgba(196,151,42,0.02)}
.panel-card-title{font-family:var(--serif);font-size:16px;font-weight:600;color:var(--gold-lt);display:inline-flex;align-items:center;gap:8px}
.panel-card-count{font-size:11px;font-weight:600;color:var(--gold);background:rgba(196,151,42,0.1);border:1px solid rgba(196,151,42,0.22);padding:3px 11px;border-radius:100px;letter-spacing:.03em}
.panel-card-body{padding:24px}
.panel-card-body.tight{padding:12px}
.panel-card-body.scroll{max-height:480px;overflow-y:auto}
.panel-card-body.scroll::-webkit-scrollbar{width:5px}
.panel-card-body.scroll::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
.panel-card-body.scroll::-webkit-scrollbar-thumb:hover{background:var(--gold-dk)}
.panel-card-desc{font-size:13px;color:var(--muted);margin-bottom:14px;line-height:1.65}
.panel-card-hint{font-size:11px;color:var(--muted);margin-top:10px;font-style:italic}
.action-buttons{display:flex;gap:8px;flex-wrap:wrap}

.admin-form{display:grid;gap:18px}
.form-group{display:flex;flex-direction:column;gap:7px}
.form-group label{font-size:11px;font-weight:600;color:var(--muted);letter-spacing:.06em;text-transform:uppercase;display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap}
.form-group label .req{color:var(--gold)}
.form-group input,.form-group textarea,.form-group select{padding:12px 14px;background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:var(--r-md);font-family:var(--body);font-size:13px;color:var(--cream);transition:var(--t);-webkit-appearance:none}
.form-group select option{background:#0f1e30;color:var(--cream)}
.form-group input::placeholder,.form-group textarea::placeholder{color:var(--muted);opacity:.6}
.form-group input:focus,.form-group textarea:focus,.form-group select:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(196,151,42,0.12),0 0 14px rgba(196,151,42,0.06);background:rgba(255,255,255,0.05)}
.form-group textarea{min-height:150px;resize:vertical;line-height:1.7;font-size:13px}
.help-text{font-size:11px;color:var(--muted);font-style:italic}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.form-actions{display:flex;gap:10px;justify-content:flex-end;align-items:center;margin-top:6px;padding-top:14px;border-top:1px solid var(--border);flex-wrap:wrap}

.admin-search-bar{padding:12px 14px;border-bottom:1px solid var(--border);background:rgba(0,0,0,0.15);position:relative}
.admin-search-bar input{width:100%;padding:8px 14px 8px 32px;background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:var(--r-sm);color:var(--cream);font-size:12px;font-family:var(--body);transition:var(--t)}
.admin-search-bar input:focus{outline:none;border-color:var(--gold);box-shadow:0 0 0 3px rgba(196,151,42,0.1)}
.admin-search-bar input::placeholder{color:var(--muted);opacity:0.7}
.admin-search-bar::before{content:'🔍';position:absolute;left:22px;top:50%;transform:translateY(-50%);font-size:11px;pointer-events:none;opacity:0.55}

.admin-song-item{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:11px 14px;background:rgba(255,255,255,0.02);border:1px solid var(--border);border-radius:var(--r-md);margin-bottom:6px;transition:var(--t)}
.admin-song-item:last-child{margin-bottom:0}
.admin-song-item:hover{border-color:var(--border-l);background:rgba(255,255,255,0.04);box-shadow:0 0 12px rgba(196,151,42,0.06)}
.admin-song-info{min-width:0;flex:1}
.admin-song-info h3{font-size:13px;font-weight:600;color:var(--cream);margin-bottom:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-family:var(--body)}
.admin-song-info p{font-size:11px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.admin-num{display:inline-flex;align-items:center;font-size:9px;font-weight:700;background:rgba(196,151,42,0.12);color:var(--gold);border:1px solid rgba(196,151,42,0.2);padding:1px 7px;border-radius:100px;margin-right:7px}
.admin-actions{display:flex;gap:5px;flex-shrink:0}
.admin-list-empty{text-align:center;color:var(--muted);padding:30px 16px;font-size:12px}

.ai-checkbox-row{margin:12px 0 4px;display:flex;align-items:center;gap:10px;padding:10px 14px;background:rgba(196,151,42,0.06);border:1px solid var(--border-l);border-radius:var(--r-md);box-shadow:0 0 12px rgba(196,151,42,0.05)}
.ai-checkbox-row input[type=checkbox]{width:auto;cursor:pointer;accent-color:var(--gold);transform:scale(1.1)}
.ai-checkbox-row label{font-size:12px;color:var(--cream);cursor:pointer;flex:1;font-weight:500}

/* ── LOGIN GATE ── */
.password-gate{max-width:420px;margin:60px auto;text-align:center;background:linear-gradient(160deg,var(--card),var(--card-2));border:1px solid var(--border-l);border-top:3px solid var(--gold);border-radius:var(--r-xl);padding:46px 42px;box-shadow:0 0 60px rgba(196,151,42,0.1),var(--shadow)}
.password-gate h2{font-family:var(--serif);font-size:26px;color:var(--cream);margin-bottom:8px}
.password-gate p{color:var(--muted);font-size:13px;margin-bottom:26px}
.password-gate .form-group{text-align:left}

/* ── TOAST ── */
.toast{position:fixed;bottom:28px;right:28px;padding:14px 22px;background:var(--card);border:1px solid var(--border-l);color:var(--cream);border-radius:var(--r-md);box-shadow:0 0 24px rgba(196,151,42,0.1),var(--shadow);font-size:13px;font-weight:500;display:flex;align-items:center;gap:9px;transform:translateY(80px);opacity:0;transition:all 0.3s cubic-bezier(.4,0,.2,1);z-index:2000;max-width:380px}
.toast.show{transform:translateY(0);opacity:1}
.toast.success{border-color:rgba(46,158,101,0.5);box-shadow:0 0 20px rgba(46,158,101,0.1),var(--shadow)}
.toast.success::before{content:'✓';color:var(--success);font-weight:700}
.toast.error{border-color:rgba(192,64,74,0.5);box-shadow:0 0 20px rgba(192,64,74,0.1),var(--shadow)}
.toast.error::before{content:'✕';color:var(--danger);font-weight:700}

.firebase-notice{background:rgba(196,151,42,0.06);border:1px solid rgba(196,151,42,0.25);border-radius:var(--r-md);padding:16px 20px;margin-bottom:24px;font-size:13px;color:var(--gold-lt);display:none}
.firebase-notice strong{color:var(--gold)}

.import-progress{margin-top:14px;background:rgba(255,255,255,0.03);border:1px solid var(--border);border-radius:var(--r-md);padding:14px 16px}
.import-progress-status{font-size:12px;color:var(--gold);margin-bottom:8px}
.import-progress-track{background:rgba(255,255,255,0.07);border-radius:4px;overflow:hidden;height:6px}
.import-progress-bar{height:100%;background:linear-gradient(90deg,var(--gold),var(--gold-lt));width:0%;transition:width 0.3s ease;box-shadow:0 0 8px rgba(196,151,42,0.4)}

/* ── PRINT ── */
@media print{
  body{background:white !important;color:#333 !important}
  .container{max-width:100%;padding:0}
  .view-section{display:block !important;border:none !important;padding:0 !important}
  .app-header,.nav-pill,.search-block,.lang-filter,.modal-actions,.btn{display:none !important}
}

/* ── RESPONSIVE ── */
@media(max-width:980px){
  .admin-layout{grid-template-columns:1fr}
  .lyrics-cols-grid{grid-template-columns:1fr}
  .lyrics-col-left{border-right:none;border-bottom:1px solid var(--border)}
}
@media(max-width:768px){
  .app-header{padding:12px 16px;flex-wrap:wrap;gap:10px}
  .header-right{order:3;width:100%}
  .quick-find{flex:1;width:auto}
  .nav-pill{order:2;width:100%;justify-content:center}
  .container{padding:24px 16px 40px}
  .page-h1{font-size:36px}
  .songs-grid{grid-template-columns:1fr}
  .form-row{grid-template-columns:1fr}
  .modal-content{padding:22px 18px;border-radius:var(--r-md)}
  .modal-song-title{font-size:26px}
  .admin-song-item{flex-direction:column;align-items:flex-start}
  .admin-actions{margin-top:8px}
  .lang-filter{flex-direction:column;align-items:stretch}
  .reset-all{margin-left:0;align-self:flex-start}
}
@media(max-width:480px){
  .nav-item{padding:7px 14px;font-size:12px}
  .modal-content{padding:18px 14px}
}
</style>
</head>
<body>

<!-- ── HEADER ── -->
<header class="app-header">
  <div class="brand-block">
    <div class="brand-icon">🎵</div>
    <div class="brand-text">
      <div class="brand-title">Choir Song Book</div>
      <div class="brand-sub">Brethren Assembly Abu Dhabi</div>
    </div>
  </div>

  <div class="nav-pill" id="mainTabs">
    <button class="nav-item on" id="viewerTab" onclick="switchTab('viewer')">Song Library</button>
    <button class="nav-item" id="adminTab" onclick="switchTab('admin')">Admin Panel</button>
  </div>

</header>

<div class="container">

  <div class="firebase-notice" id="firebaseNotice">
    <strong>⚙ Firebase Setup Required:</strong> Please replace the placeholder Firebase config below with your project credentials.
  </div>

  <!-- ══════════ SONG LIBRARY ══════════ -->
  <div id="viewer" class="view-section active">

    <h1 class="page-h1">Song Library</h1>

    <div class="search-block">
      <div class="search-block-label">Search the Library</div>
      <div class="search-row">
        <div class="search-input-wrap">
          <span class="search-icon">🔍</span>
          <input type="text" class="search-input" id="searchInput"
                 placeholder="Search by title, artist, lyrics, language, or serial #…"
                 oninput="onSearchInput()">
      </div>
    </div>

    <div class="lang-filter" id="languageFilter">
      <span class="lang-chip active" onclick="filterLanguage('all')">All Languages</span>
      <button class="reset-all" id="resetAllBtn" onclick="resetFilters()" style="display:none">RESET ALL ✕</button>
    </div>

    <div id="songsContainer">
      <div class="loading-state"><div class="spinner"></div><p style="color:var(--muted);font-size:13px">Loading songs…</p></div>
    </div>
  </div>

  <!-- ══════════ ADMIN PANEL ══════════ -->
  <div id="admin" class="view-section">

    <div id="loginSection" class="password-gate">
      <h2>Admin Access</h2>
      <p>Enter your admin password to manage the song library</p>
      <div class="form-group" style="margin-bottom:16px">
        <label>Password</label>
        <input type="password" id="superAdminPassword" placeholder="Enter password…"
               onkeypress="if(event.key==='Enter')loginAdmin()">
      </div>
      <button class="btn btn-primary" style="width:100%" onclick="loginAdmin()">Sign In</button>
    </div>

    <div id="adminPanel" style="display:none">

      <div class="page-tag">Management Dashboard</div>
      <div class="admin-page-head">
        <h1 class="page-h1" style="font-size:36px;margin-bottom:0">Brethren Assembly · Choir</h1>
        <div class="admin-actions-top">
          <button class="btn btn-secondary btn-sm" onclick="exportSongs()">⬇ Backup JSON</button>
          <button class="btn btn-primary btn-sm" onclick="document.getElementById('importExcelFile').click()">📥 Excel Import</button>
          <input type="file" id="importExcelFile" accept=".xlsx,.xls" style="display:none" onchange="importFromExcel(this)">
          <button class="btn btn-secondary btn-sm" onclick="logoutAdmin()">Sign Out</button>
        </div>
      </div>

      <div class="stats-grid" id="statsSection"></div>

      <div class="admin-layout">

        <div class="panel-card">
          <div class="panel-card-head">
            <span class="panel-card-title" id="formTitle">✦ Add New Song</span>
            <button type="button" class="btn btn-danger btn-sm" id="cancelEditBtn" style="display:none" onclick="clearForm()">Cancel Edit</button>
          </div>
          <div class="panel-card-body">

            <form class="admin-form" onsubmit="saveSong(event)">
              <div class="form-row">
                <div class="form-group">
                  <label>Title <span class="req">*</span></label>
                  <input type="text" id="songTitle" required placeholder="e.g. Amazing Grace">
                </div>
                <div class="form-group">
                  <label>Artist / Band</label>
                  <input type="text" id="songArtist" placeholder="e.g. Traditional">
                </div>
              </div>
              <div class="form-row">
                <div class="form-group">
                  <label>Language <span class="req">*</span></label>
                  <select id="songLanguage" required onchange="toggleTransliterationFields()">
                    <option value="">Select language…</option>
                    <option value="English">English</option>
                    <option value="Hindi">Hindi</option>
                    <option value="Tamil">Tamil</option>
                    <option value="Telugu">Telugu</option>
                    <option value="Malayalam">Malayalam</option>
                    <option value="Kannada">Kannada</option>
                    <option value="Bengali">Bengali</option>
                    <option value="Marathi">Marathi</option>
                    <option value="Gujarati">Gujarati</option>
                    <option value="Punjabi">Punjabi</option>
                    <option value="Urdu">Urdu</option>
                    <option value="Spanish">Spanish</option>
                    <option value="French">French</option>
                    <option value="Other">Other</option>
                  </select>
                </div>
                <div class="form-group">
                  <label>Key / Tempo</label>
                  <input type="text" id="songKey" placeholder="e.g. Key of G, 120 BPM">
                </div>
              </div>
              <div class="form-group">
                <label>YouTube Link</label>
                <input type="url" id="songYoutube" placeholder="https://www.youtube.com/watch?v=…">
                <span class="help-text">Paste a YouTube URL — it will display as an embedded player</span>
              </div>

              <div id="transliterationSection" style="display:none;gap:14px">
                <div class="form-group">
                  <label>Original Language Lyrics <span class="req">*</span></label>
                  <textarea id="songLyricsOriginal" placeholder="Paste lyrics in the original language…"></textarea>
                  <span class="help-text">Use [Verse 1], [Chorus] etc. as section headers</span>
                </div>
                <div class="form-group">
                  <label>
                    <span>English Transliteration <span class="req">*</span></span>
                  </label>
                  <textarea id="songLyricsTransliteration" placeholder="Paste phonetic English spelling…"></textarea>
                  <span class="help-text">Match lines exactly to the original for side-by-side display</span>
                </div>
              </div>

              <div id="englishLyricsSection">
                <div class="form-group">
                  <label>Lyrics <span class="req">*</span></label>
                  <textarea id="songLyricsEnglish" required placeholder="Paste song lyrics…"></textarea>
                </div>
              </div>

              <div class="form-actions">
                <button type="button" class="btn btn-secondary" onclick="clearForm()">Clear</button>
                <button type="submit" class="btn btn-primary" id="submitBtn">💾 Save Song</button>
              </div>
            </form>
          </div>
        </div>

        <div class="admin-side-stack">

          <div class="panel-card">
            <div class="panel-card-head">
              <span class="panel-card-title">📋 Song Management</span>
              <span class="panel-card-count" id="adminSongCount">0 songs</span>
            </div>
            <div class="admin-search-bar">
              <input type="text" id="adminSearch" placeholder="Search by title, artist, or language…" oninput="renderAdminList()">
            </div>
            <div class="panel-card-body tight scroll">
              <div id="adminSongList"></div>
            </div>
          </div>

          <div class="panel-card">
            <div class="panel-card-head">
              <span class="panel-card-title">📥 Bulk Import</span>
            </div>
            <div class="panel-card-body">
              <p class="panel-card-desc">Upload the Excel template to add many songs at once. Supports English and non-English templates.</p>
              <div class="ai-checkbox-row">
                <input type="checkbox" id="cleanWithAI">
                <label for="cleanWithAI">✨ Clean lyrics formatting with AI before importing</label>
              </div>
              <div class="action-buttons">
                <button class="btn btn-primary btn-sm" onclick="document.getElementById('importExcelFile').click()">📊 Upload Template</button>
              </div>
              <p class="panel-card-hint">Duplicate songs are skipped automatically.</p>
              <div id="excelImportProgress" style="display:none" class="import-progress">
                <div id="excelImportStatus" class="import-progress-status">Processing...</div>
                <div class="import-progress-track">
                  <div id="excelProgressBar" class="import-progress-bar"></div>
                </div>
              </div>
            </div>
          </div>

          <div class="panel-card">
            <div class="panel-card-head">
              <span class="panel-card-title">💾 Backup &amp; Restore</span>
            </div>
            <div class="panel-card-body">
              <p class="panel-card-desc">Download all songs as JSON, or restore from a previous backup file.</p>
              <div class="action-buttons">
                <button class="btn btn-secondary btn-sm" onclick="exportSongs()">⬇ Export JSON</button>
                <button class="btn btn-secondary btn-sm" onclick="document.getElementById('importFile').click()">⬆ Import JSON</button>
                <input type="file" id="importFile" accept=".json" style="display:none" onchange="importSongs(this)">
              </div>
            </div>
          </div>

        </div>
      </div>

    </div>
  </div>

</div>

<!-- ── LYRICS MODAL ── -->
<div class="modal" id="lyricsModal" onclick="if(event.target===this)closeModal()">
  <div class="modal-content">
    <button class="close-btn" onclick="closeModal()">✕</button>
    <div id="modalContent"></div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
// ════════════════════════════════════════════════
//  FIREBASE CONFIG
// ════════════════════════════════════════════════
const firebaseConfig = {
  apiKey: "AIzaSyBwJ_WRRC1REMOu-Ybtu53I5xKX9t5AkHw",
  authDomain: "choir-song-book.firebaseapp.com",
  projectId: "choir-song-book",
  storageBucket: "choir-song-book.firebasestorage.app",
  messagingSenderId: "1079381710164",
  appId: "1:1079381710164:web:fb6ceac8b37f59c166a4c1",
  measurementId: "G-W6VNJPMDS2"
};

const AI_PROXY_URL = '/.netlify/functions/gemini';

let db = null;
let firebaseReady = false;

function initFirebase() {
  try {
    firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();
    firebaseReady = true;
    setStatus('live', 'Live');
    document.getElementById('firebaseNotice').style.display = 'none';
    startRealtimeListener();
  } catch(err) {
    console.warn('Firebase init error:', err);
    setStatus('', 'Offline');
    document.getElementById('firebaseNotice').style.display = 'block';
    fallbackToLocalStorage();
  }
}

function setStatus(type, text) {
  var dot = document.getElementById('statusDot');
  var txt = document.getElementById('statusText');
  if(dot) dot.className = 'status-dot' + (type ? ' ' + type : '');
  if(txt) txt.textContent = text;
}

if(firebaseConfig.apiKey === 'YOUR_API_KEY') {
  document.getElementById('firebaseNotice').style.display = 'block';
  setStatus('', 'Offline');
  fallbackToLocalStorage();
} else {
  initFirebase();
}

const SUPER_ADMIN_PASSWORD_HASH = 'bd21dd361c7e82747d837ba2557abef9d170c2d21d070ad78b072def0dd36208';
async function hashPassword(pw) {
  var buf = new TextEncoder().encode(pw);
  var hash = await crypto.subtle.digest('SHA-256', buf);
  return Array.from(new Uint8Array(hash)).map(function(b){ return b.toString(16).padStart(2,'0'); }).join('');
}
var songs = [];
var currentLanguage = 'all';
var editingId = null;
var currentUser = null;
var unsubscribeListener = null;

// ════════════════════════════════════════════════
//  AI PROXY CLIENT
// ════════════════════════════════════════════════
async function callAI(systemPrompt, userPrompt, maxTokens) {
  maxTokens = maxTokens || 2000;
  var res = await fetch(AI_PROXY_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ system: systemPrompt, userPrompt: userPrompt, maxTokens: maxTokens })
  });
  var data = await res.json();
  if(!res.ok) throw new Error(data.error || 'AI request failed');
  return data.text || '';
}

function parseAIJson(text) {
  var cleaned = text.replace(/```(?:json)?/gi, '').trim();
  var positions = [cleaned.indexOf('{'), cleaned.indexOf('[')].filter(function(i){ return i >= 0; });
  var start = positions.length ? Math.min.apply(null, positions) : -1;
  var end = Math.max(cleaned.lastIndexOf('}'), cleaned.lastIndexOf(']'));
  if(start >= 0 && end > start) cleaned = cleaned.slice(start, end+1);
  return JSON.parse(cleaned);
}

// ════════════════════════════════════════════════
//  SEARCH
// ════════════════════════════════════════════════
function onSearchInput() {
  searchSongs();
}

function onQuickFind(e) {
  if(e.key === 'Enter' || e.key === 'Return') {
    var q = document.getElementById('quickFind').value.trim();
    if(!q) return;
    switchTab('viewer');
    document.getElementById('searchInput').value = q;
    document.getElementById('quickFind').value = '';
    searchSongs();
    document.getElementById('searchInput').focus();
  }
}

// ════════════════════════════════════════════════
//  BULK CLEAN
// ════════════════════════════════════════════════
async function cleanSongWithAI(song) {
  var isEn = song.language === 'English';
  var system = 'Clean up song lyrics for a choir database. Rules: (1) Fix capitalization. (2) Add section headers like [Verse 1], [Chorus], [Bridge] where lyrics structure makes them obvious. (3) Fix spacing/formatting. (4) Preserve ALL original words — do not add or remove content. Output ONLY the cleaned lyrics. No preamble, no markdown.';
  try {
    if(isEn && song.lyricsEnglish) song.lyricsEnglish = await callAI(system, song.lyricsEnglish, 3000);
    else {
      if(song.lyricsOriginal) song.lyricsOriginal = await callAI(system, song.lyricsOriginal, 3000);
      if(song.lyricsTransliteration) song.lyricsTransliteration = await callAI(system, song.lyricsTransliteration, 3000);
    }
  } catch(err) {
    console.warn('AI clean failed for', song.title, err.message);
  }
  return song;
}

// ════════════════════════════════════════════════
//  CORE APP
// ════════════════════════════════════════════════
function startRealtimeListener() {
  if(!db) return;
  if(unsubscribeListener) unsubscribeListener();
  unsubscribeListener = db.collection('songs').onSnapshot(function(snapshot) {
    songs = [];
    snapshot.forEach(function(doc){ songs.push(Object.assign({ id: doc.id }, doc.data())); });
    cacheSongsLocally(songs);
    renderSongs(); updateLanguages();
    hideOfflineBanner();
    if(document.getElementById('adminPanel').style.display !== 'none') { updateStats(); renderAdminList(); }
  }, function(err) { console.error('Firestore error:', err); loadFromCache(); });
}

function fallbackToLocalStorage() {
  songs = JSON.parse(localStorage.getItem('songsCollection') || '[]');
  renderSongs(); updateLanguages(); checkLoginStatus();
}
function saveToLocalStorage() { localStorage.setItem('songsCollection', JSON.stringify(songs)); }

document.addEventListener('DOMContentLoaded', function() { checkLoginStatus(); });

function checkLoginStatus() {
  var saved = sessionStorage.getItem('currentUser');
  if(saved) { currentUser = JSON.parse(saved); showAdminDashboard(); }
}

async function loginAdmin() {
  var pw = document.getElementById('superAdminPassword').value;
  var hashed = await hashPassword(pw);
  if(hashed === SUPER_ADMIN_PASSWORD_HASH) {
    currentUser = { type:'super', username:'Super Admin' };
    sessionStorage.setItem('currentUser', JSON.stringify(currentUser));
    showAdminDashboard();
    showToast('Welcome, Super Admin!', 'success');
  } else {
    showToast('Incorrect password', 'error');
  }
}

function showAdminDashboard() {
  document.getElementById('loginSection').style.display = 'none';
  document.getElementById('adminPanel').style.display = 'block';
  updateStats(); renderAdminList();
}

function logoutAdmin() {
  currentUser = null;
  sessionStorage.removeItem('currentUser');
  document.getElementById('adminPanel').style.display = 'none';
  document.getElementById('loginSection').style.display = 'block';
  switchTab('viewer');
  showToast('Signed out', 'success');
}

function switchTab(tab) {
  document.querySelectorAll('.nav-item').forEach(function(t){ t.classList.remove('on'); });
  document.querySelectorAll('.view-section').forEach(function(s){ s.classList.remove('active'); });
  document.getElementById(tab + 'Tab').classList.add('on');
  document.getElementById(tab).classList.add('active');
  if(tab === 'admin') { updateStats(); renderAdminList(); }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function computeSerialNumbers(songsList) {
  var groups = {};
  songsList.forEach(function(s){ var l = s.language || 'Unknown'; if(!groups[l]) groups[l] = []; groups[l].push(s); });
  var map = {};
  Object.keys(groups).forEach(function(l){
    groups[l].sort(function(a,b){ return (a.title||'').toLowerCase().localeCompare((b.title||'').toLowerCase()); });
    groups[l].forEach(function(s,i){ map[s.id] = i+1; });
  });
  return map;
}

function renderSongs(filteredSongs) {
  var container = document.getElementById('songsContainer');
  var list = filteredSongs !== undefined ? filteredSongs : songs;

  if(list.length === 0) {
    container.innerHTML = '<div class="empty-state"><div class="empty-icon">🎶</div><h3>' + (filteredSongs !== undefined ? 'No songs match your search' : 'No songs yet') + '</h3><p>' + (filteredSongs !== undefined ? 'Try different keywords or reset filters' : 'Use the Admin Panel to add your first song') + '</p></div>';
    return;
  }

  var serialMap = computeSerialNumbers(list);
  var grouped = {};
  list.forEach(function(s){ var l = s.language || 'Unknown'; if(!grouped[l]) grouped[l] = []; grouped[l].push(s); });

  var html = '';
  Object.keys(grouped).sort().forEach(function(lang){
    var langSongs = grouped[lang].sort(function(a,b){ return (a.title||'').toLowerCase().localeCompare((b.title||'').toLowerCase()); });
    html += '<div class="lang-group"><div class="lang-group-head"><span class="lang-group-title">' + escHtml(lang) + '</span><span class="lang-group-count">' + langSongs.length + ' song' + (langSongs.length!==1?'s':'') + '</span></div><div class="songs-grid">';
    langSongs.forEach(function(song){
      var serial = serialMap[song.id];
      var isEn = song.language === 'English';
      var hasYt = song.youtube && song.youtube.trim();
      var ytId = hasYt ? extractYTId(song.youtube) : null;
      html += '<div class="song-card" onclick="showLyrics(\'' + song.id + '\')">' +
        '<div class="song-card-tags">' +
          '<span class="song-tag serial">#' + serial + '</span>' +
          '<span class="song-tag ' + (isEn ? 'english' : '') + '">' + escHtml(song.language||'Other') + '</span>' +
          (ytId ? '<span class="song-tag video">▶ Video</span>' : '') +
        '</div>' +
        '<div class="song-card-title">' + escHtml(song.title) + '</div>' +
        '<div class="song-card-artist">' + escHtml(song.artist||'Unknown') + '</div>' +
        '<div class="song-card-foot">' +
          '<span class="song-card-meta">' + (song.key ? '♩ ' + escHtml(song.key) : 'View lyrics') + '</span>' +
          '<span class="song-card-arrow">→</span>' +
        '</div>' +
      '</div>';
    });
    html += '</div></div>';
  });
  container.innerHTML = html;
}

function updateLanguages() {
  var langs = [];
  songs.forEach(function(s){ if(s.language && langs.indexOf(s.language) < 0) langs.push(s.language); });
  var html = '<span class="lang-chip' + (currentLanguage === 'all' ? ' active' : '') + '" onclick="filterLanguage(\'all\')">All Languages</span>';
  langs.sort().forEach(function(l){
    html += '<span class="lang-chip' + (currentLanguage === l ? ' active' : '') + '" onclick="filterLanguage(\'' + escHtml(l) + '\')">' + escHtml(l) + '</span>';
  });
  html += '<button class="reset-all" id="resetAllBtn" onclick="resetFilters()" style="' + ((currentLanguage !== 'all' || document.getElementById('searchInput').value) ? '' : 'display:none') + '">RESET ALL ✕</button>';
  document.getElementById('languageFilter').innerHTML = html;
}

function filterLanguage(lang) {
  currentLanguage = lang;
  updateLanguages();
  searchSongs();
}

function resetFilters() {
  currentLanguage = 'all';
  document.getElementById('searchInput').value = '';
  updateLanguages();
  renderSongs();
}

function searchSongs() {
  var q = document.getElementById('searchInput').value.toLowerCase().trim();
  if(!q && currentLanguage === 'all') { renderSongs(); updateLanguages(); return; }

  var isNum = /^\d+$/.test(q);
  var serialMap = computeSerialNumbers(songs);
  var filtered = songs.filter(function(s){
    var matchLang = currentLanguage === 'all' || s.language === currentLanguage;
    if(!matchLang) return false;
    if(!q) return true;
    if(isNum) return serialMap[s.id] === parseInt(q);
    return (s.title && s.title.toLowerCase().includes(q)) ||
           (s.artist && s.artist.toLowerCase().includes(q)) ||
           (s.language && s.language.toLowerCase().includes(q)) ||
           (s.lyricsEnglish && s.lyricsEnglish.toLowerCase().includes(q)) ||
           (s.lyricsOriginal && s.lyricsOriginal.toLowerCase().includes(q)) ||
           (s.lyricsTransliteration && s.lyricsTransliteration.toLowerCase().includes(q));
  });
  renderSongs(filtered);
  updateLanguages();
}

// ════════════════════════════════════════════════
//  LYRICS MODAL
// ════════════════════════════════════════════════
function showLyrics(id) {
  var song = songs.find(function(s){ return s.id === id; });
  if(!song) return;
  var isEn = song.language === 'English';
  var hasYt = song.youtube && song.youtube.trim();
  var ytId = hasYt ? extractYTId(song.youtube) : null;

  var lyricsBlock = '';
  if(isEn) {
    lyricsBlock = '<div class="lyrics-display">' + escHtml(song.lyricsEnglish || song.lyrics || '') + '</div>';
  } else {
    var origLines = (song.lyricsOriginal || '').split('\n');
    var transLines = (song.lyricsTransliteration || '').split('\n');
    var maxLines = Math.max(origLines.length, transLines.length);
    var leftHtml = '', rightHtml = '';
    for(var i = 0; i < maxLines; i++) {
      var orig = origLines[i] || '', trans = transLines[i] || '';
      leftHtml += escHtml(orig) + '\n';
      rightHtml += escHtml(trans) + '\n';
    }
    lyricsBlock = '<div class="lyrics-table-head"><span class="lyrics-table-label">' + escHtml(song.language) + ' with English transliteration</span>' +
      '<div class="toggle-view"><button class="toggle-btn active" onclick="toggleLyricsView(\'side\')" id="btnSide">Side by Side</button><button class="toggle-btn" onclick="toggleLyricsView(\'orig\')" id="btnOrig">Original Only</button><button class="toggle-btn" onclick="toggleLyricsView(\'trans\')" id="btnTrans">English Only</button></div></div>' +
      '<div id="lyricsViewWrap">' +
        '<div class="lyrics-cols-grid" id="tblSide">' +
          '<div class="lyrics-col lyrics-col-left"><div class="lyrics-col-label">' + escHtml(song.language) + ' Original</div><div class="lyrics-col-content">' + leftHtml + '</div></div>' +
          '<div class="lyrics-col lyrics-col-right"><div class="lyrics-col-label">English Transliteration</div><div class="lyrics-col-content">' + rightHtml + '</div></div>' +
        '</div>' +
        '<div id="viewOrig" style="display:none" class="lyrics-display">' + escHtml(song.lyricsOriginal || '') + '</div>' +
        '<div id="viewTrans" style="display:none" class="lyrics-display">' + escHtml(song.lyricsTransliteration || '') + '</div>' +
      '</div>';
  }

  var refBlock = '';
  if(hasYt && ytId) {
    refBlock = '<div class="ref-material">' +
      '<div class="ref-material-label">Reference Material</div>' +
      '<div class="ref-card">' +
        '<div class="ref-card-video"><iframe src="https://www.youtube.com/embed/' + ytId + '" title="YouTube video" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></div>' +
        '<div class="ref-card-body">Performance · YouTube</div>' +
      '</div></div>';
  } else if(hasYt) {
    refBlock = '<div class="ref-material"><div class="ref-material-label">Reference Material</div><a href="' + escHtml(song.youtube) + '" target="_blank" style="color:#f27680;font-size:13px;text-decoration:underline">▶ Watch on YouTube</a></div>';
  }

  var serialMap = computeSerialNumbers(songs);
  var serial = serialMap[song.id];

  document.getElementById('modalContent').innerHTML =
    '<div class="modal-back-row">' +
      '<button class="modal-back" onclick="closeModal()">← Back to Library</button>' +
      '<div class="modal-actions">' +
        '<button class="btn btn-secondary btn-sm" onclick="printLyrics(\'' + song.id + '\')">🖨 Print Lyrics</button>' +
        '<button class="btn btn-secondary btn-sm" onclick="copyLyrics(\'' + song.id + '\')">📋 Copy Text</button>' +
      '</div>' +
    '</div>' +
    '<div class="modal-tags-row">' +
      '<span class="song-tag serial">#' + serial + '</span>' +
      '<span class="song-tag ' + (isEn ? 'english' : '') + '">' + escHtml(song.language||'Other') + '</span>' +
      (song.key ? '<span class="song-tag">♩ ' + escHtml(song.key) + '</span>' : '') +
      (hasYt ? '<span class="song-tag video">▶ Video</span>' : '') +
    '</div>' +
    '<h1 class="modal-song-title">' + escHtml(song.title) + '</h1>' +
    '<div class="modal-song-artist">' + escHtml(song.artist||'Unknown') + '</div>' +
    lyricsBlock + refBlock;

  document.getElementById('lyricsModal').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function toggleLyricsView(view) {
  ['btnSide','btnOrig','btnTrans'].forEach(function(id){ document.getElementById(id) && document.getElementById(id).classList.remove('active'); });
  ['tblSide','viewOrig','viewTrans'].forEach(function(id){ var el = document.getElementById(id); if(el) el.style.display = 'none'; });
  if(view === 'side') { document.getElementById('btnSide').classList.add('active'); document.getElementById('tblSide').style.display = 'grid'; }
  else if(view === 'orig') { document.getElementById('btnOrig').classList.add('active'); document.getElementById('viewOrig').style.display = 'block'; }
  else { document.getElementById('btnTrans').classList.add('active'); document.getElementById('viewTrans').style.display = 'block'; }
}

function closeModal() {
  document.getElementById('lyricsModal').classList.remove('active');
  document.body.style.overflow = '';
}

// ════════════════════════════════════════════════
//  ADMIN
// ════════════════════════════════════════════════
function updateStats() {
  var langs = [];
  songs.forEach(function(s){ if(s.language && langs.indexOf(s.language) < 0) langs.push(s.language); });

  // New this month
  var now = new Date();
  var thisMonth = now.getMonth(), thisYear = now.getFullYear();
  var newThisMonth = songs.filter(function(s){
    if(!s.dateAdded) return false;
    var d = new Date(s.dateAdded);
    return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
  }).length;

  // Active setlists — placeholder, would need real persistence
  var activeSetlists = 0;

  document.getElementById('statsSection').innerHTML =
    '<div class="stat-card"><div class="stat-label">Total Songs</div><div class="stat-number">' + String(songs.length).padStart(3, '0') + '</div></div>' +
    '<div class="stat-card"><div class="stat-label">Languages</div><div class="stat-number">' + String(langs.length).padStart(2, '0') + '</div></div>' +
    '<div class="stat-card"><div class="stat-label">Active Setlists</div><div class="stat-number">' + String(activeSetlists).padStart(2, '0') + '</div></div>' +
    '<div class="stat-card"><div class="stat-label">New This Month</div><div class="stat-number positive">+' + newThisMonth + '</div></div>';
}

function toggleTransliterationFields() {
  var lang = document.getElementById('songLanguage').value;
  var isEn = lang === 'English' || !lang;
  document.getElementById('transliterationSection').style.display = isEn ? 'none' : 'grid';
  document.getElementById('englishLyricsSection').style.display = isEn ? 'block' : 'none';
  document.getElementById('songLyricsOriginal').required = !isEn;
  document.getElementById('songLyricsTransliteration').required = !isEn;
  document.getElementById('songLyricsEnglish').required = isEn;
}

async function saveSong(e) {
  e.preventDefault();
  var lang = document.getElementById('songLanguage').value;
  var isEn = lang === 'English';
  var songData = {
    title:    document.getElementById('songTitle').value.trim(),
    artist:   document.getElementById('songArtist').value.trim(),
    language: lang,
    key:      document.getElementById('songKey').value.trim(),
    youtube:  document.getElementById('songYoutube').value.trim(),
    dateAdded: editingId
      ? (songs.find(function(s){ return s.id === editingId; }) || {}).dateAdded || new Date().toISOString()
      : new Date().toISOString()
  };
  if(isEn) songData.lyricsEnglish = document.getElementById('songLyricsEnglish').value.trim();
  else { songData.lyricsOriginal = document.getElementById('songLyricsOriginal').value.trim(); songData.lyricsTransliteration = document.getElementById('songLyricsTransliteration').value.trim(); }

  try {
    if(firebaseReady && db) {
      if(editingId) await db.collection('songs').doc(editingId).update(songData);
      else await db.collection('songs').add(songData);
    } else {
      if(editingId) {
        var idx = songs.findIndex(function(s){ return s.id === editingId; });
        if(idx >= 0) songs[idx] = Object.assign({ id: editingId }, songData);
      } else songs.push(Object.assign({ id: Date.now().toString() }, songData));
      saveToLocalStorage();
      renderSongs(); updateLanguages(); updateStats(); renderAdminList();
    }
    showToast(editingId ? 'Song updated!' : 'Song added!', 'success');
    clearForm();
  } catch(err) {
    showToast('Save failed: ' + err.message, 'error');
  }
}

function clearForm() {
  editingId = null;
  ['songTitle','songArtist','songKey','songYoutube','songLyricsEnglish','songLyricsOriginal','songLyricsTransliteration'].forEach(function(id){
    document.getElementById(id).value = '';
  });
  document.getElementById('songLanguage').value = '';
  document.getElementById('transliterationSection').style.display = 'none';
  document.getElementById('englishLyricsSection').style.display = 'block';
  document.getElementById('submitBtn').textContent = '💾 Save Song';
  document.getElementById('cancelEditBtn').style.display = 'none';
  document.getElementById('formTitle').textContent = '✦ Add New Song';
}

function editSong(id) {
  var song = songs.find(function(s){ return s.id === id; });
  if(!song) return;
  editingId = id;
  document.getElementById('songTitle').value = song.title || '';
  document.getElementById('songArtist').value = song.artist || '';
  document.getElementById('songLanguage').value = song.language || '';
  document.getElementById('songKey').value = song.key || '';
  document.getElementById('songYoutube').value = song.youtube || '';
  toggleTransliterationFields();
  if(song.language === 'English') document.getElementById('songLyricsEnglish').value = song.lyricsEnglish || song.lyrics || '';
  else {
    document.getElementById('songLyricsOriginal').value = song.lyricsOriginal || '';
    document.getElementById('songLyricsTransliteration').value = song.lyricsTransliteration || '';
  }
  document.getElementById('submitBtn').textContent = '💾 Update Song';
  document.getElementById('cancelEditBtn').style.display = 'inline-flex';
  document.getElementById('formTitle').textContent = '✎ Edit Song';
  window.scrollTo({ top: 0, behavior: 'smooth' });
  showToast('Editing: ' + song.title, 'success');
}

async function deleteSong(id) {
  if(!confirm('Delete this song? This cannot be undone.')) return;
  try {
    if(firebaseReady && db) await db.collection('songs').doc(id).delete();
    else {
      songs = songs.filter(function(s){ return s.id !== id; });
      saveToLocalStorage();
      renderSongs(); updateLanguages(); updateStats(); renderAdminList();
    }
    if(editingId === id) clearForm();
    showToast('Song deleted', 'success');
  } catch(err) {
    showToast('Delete failed: ' + err.message, 'error');
  }
}

function renderAdminList() {
  var container = document.getElementById('adminSongList');
  var countEl = document.getElementById('adminSongCount');
  var searchEl = document.getElementById('adminSearch');
  var query = searchEl ? searchEl.value.toLowerCase().trim() : '';
  var visible = songs;
  if(query) {
    visible = songs.filter(function(s){
      return (s.title && s.title.toLowerCase().includes(query)) ||
             (s.artist && s.artist.toLowerCase().includes(query)) ||
             (s.language && s.language.toLowerCase().includes(query));
    });
  }
  if(countEl) countEl.textContent = query ? visible.length + ' of ' + songs.length : songs.length + ' song' + (songs.length !== 1 ? 's' : '');

  if(songs.length === 0) { container.innerHTML = '<div class="admin-list-empty">No songs added yet</div>'; return; }
  if(visible.length === 0) { container.innerHTML = '<div class="admin-list-empty">No songs match &ldquo;' + escHtml(query) + '&rdquo;</div>'; return; }

  var serialMap = computeSerialNumbers(songs);
  var sorted = visible.slice().sort(function(a,b){ return new Date(b.dateAdded) - new Date(a.dateAdded); });
  var html = '';
  sorted.forEach(function(song){
    html += '<div class="admin-song-item">' +
      '<div class="admin-song-info"><h3><span class="admin-num">#' + serialMap[song.id] + '</span>' + escHtml(song.title) + '</h3>' +
      '<p>' + escHtml(song.artist) + ' · ' + escHtml(song.language || 'Unknown') + (song.language !== 'English' ? ' · transliteration' : '') + '</p></div>' +
      '<div class="admin-actions">' +
        '<button class="btn btn-secondary btn-sm" onclick="editSong(\'' + song.id + '\')">Edit</button>' +
        '<button class="btn btn-danger btn-sm" onclick="deleteSong(\'' + song.id + '\')">Delete</button>' +
      '</div>' +
    '</div>';
  });
  container.innerHTML = html;
}

function copyLyrics(id) {
  var song = songs.find(function(s){ return s.id === id; });
  if(!song) return;
  var txt = song.title + '\nby ' + song.artist + '\n\n';
  if(song.language === 'English') txt += song.lyricsEnglish || song.lyrics || '';
  else txt += song.language + ' (Original):\n' + (song.lyricsOriginal || '') + '\n\nEnglish Transliteration:\n' + (song.lyricsTransliteration || '');
  navigator.clipboard.writeText(txt).then(function(){ showToast('Lyrics copied!', 'success'); });
}

function printLyrics(id) {
  var song = songs.find(function(s){ return s.id === id; });
  if(!song) return;
  var lyricsContent;
  if(song.language === 'English') {
    lyricsContent = '<div class="lyrics-display">' + escHtml(song.lyricsEnglish || song.lyrics || '') + '</div>';
  } else {
    var origLines = (song.lyricsOriginal||'').split('\n');
    var transLines = (song.lyricsTransliteration||'').split('\n');
    var rows = '';
    for(var i = 0; i < Math.max(origLines.length, transLines.length); i++) {
      var o = origLines[i]||'', t = transLines[i]||'';
      if(!o.trim()&&!t.trim()) continue;
      if(o.trim().startsWith('[')&&o.trim().endsWith(']')) rows += '<tr style="background:#eee"><td colspan="2" style="text-align:center;font-weight:700;padding:8px">' + escHtml(o.trim()) + '</td></tr>';
      else rows += '<tr><td style="width:50%;padding:10px 14px;border-bottom:1px solid #ddd;white-space:pre-wrap">' + escHtml(o) + '</td><td style="width:50%;padding:10px 14px;border-bottom:1px solid #ddd;color:#555;white-space:pre-wrap">' + escHtml(t) + '</td></tr>';
    }
    lyricsContent = '<table style="width:100%;border-collapse:collapse"><thead><tr><th style="background:#333;color:white;padding:10px 14px;text-align:left">' + escHtml(song.language) + '</th><th style="background:#333;color:white;padding:10px 14px;text-align:left">English Transliteration</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }
  var pw = window.open('', '_blank');
  pw.document.write('<!DOCTYPE html><html><head><title>' + escHtml(song.title) + '</title><style>body{font-family:Georgia,serif;padding:40px;max-width:900px;margin:0 auto;color:#333}.h{text-align:center;margin-bottom:30px;border-bottom:2px solid #333;padding-bottom:20px}h1{font-size:1.8em;margin-bottom:8px}p{color:#666;margin:4px 0}.lyrics-display{white-space:pre-wrap;line-height:2;font-size:1.1em;padding:20px;background:#fafafa;border-radius:8px;margin-top:20px}@media print{body{padding:20px}}</style></head><body><div class="h"><h1>' + escHtml(song.title) + '</h1><p>' + escHtml(song.artist) + ' · ' + escHtml(song.language||'Unknown') + '</p>' + (song.key?'<p>'+escHtml(song.key)+'</p>':'') + '<p style="font-size:0.85em;margin-top:12px;color:#999">Brethren Assembly Abu Dhabi — Choir Song Book</p></div>' + lyricsContent + '</body></html>');
  pw.document.close();
  setTimeout(function(){ pw.print(); }, 400);
}

function exportSongs() {
  var blob = new Blob([JSON.stringify(songs, null, 2)], { type:'application/json' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = 'songs-backup-' + new Date().toISOString().split('T')[0] + '.json'; a.click();
  URL.revokeObjectURL(url);
  showToast('Exported!', 'success');
}

async function importSongs(input) {
  var file = input.files[0];
  if(!file) return;
  var reader = new FileReader();
  reader.onload = async function(e) {
    try {
      var imported = JSON.parse(e.target.result);
      if(!Array.isArray(imported)) throw new Error('Invalid format');
      var newCount = 0;
      var existingIds = songs.map(function(s){ return s.id; });
      if(firebaseReady && db) {
        var batch = db.batch();
        imported.forEach(function(s){
          if(!existingIds.includes(s.id)) {
            var ref = db.collection('songs').doc(s.id || Date.now().toString() + Math.random());
            var data = Object.assign({}, s); delete data.id;
            batch.set(ref, data);
            newCount++;
          }
        });
        await batch.commit();
      } else {
        imported.forEach(function(s){ if(!existingIds.includes(s.id)) { songs.push(s); newCount++; } });
        saveToLocalStorage(); renderSongs(); updateLanguages(); updateStats(); renderAdminList();
      }
      showToast('Imported ' + newCount + ' new songs!', 'success');
    } catch(err) { showToast('Import error: ' + err.message, 'error'); }
  };
  reader.readAsText(file);
  input.value = '';
}

function extractYTId(url) {
  if(!url) return null;
  var m = url.match(/^.*(youtu\.be\/|v\/|embed\/|watch\?v=|&v=)([^#&?]*).*/);
  return (m && m[2].length === 11) ? m[2] : null;
}

function escHtml(text) {
  if(!text) return '';
  var d = document.createElement('div'); d.textContent = text; return d.innerHTML;
}

function showToast(msg, type) {
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast ' + (type || 'success') + ' show';
  setTimeout(function(){ t.classList.remove('show'); }, 3500);
}

document.addEventListener('keydown', function(e){
  if(e.key === 'Escape') { closeModal(); closeAutofillModal(); }
});

// ─── Offline Support ─────────────────────────────
var CACHE_KEY = 'songsOfflineCache';
var offlineBannerShown = false;
function cacheSongsLocally(songsList) {
  try { localStorage.setItem(CACHE_KEY, JSON.stringify(songsList)); localStorage.setItem(CACHE_KEY + '_time', new Date().toISOString()); } catch(e) {}
}
function loadFromCache() {
  var cached = localStorage.getItem(CACHE_KEY);
  if(cached) {
    songs = JSON.parse(cached);
    var cacheTime = localStorage.getItem(CACHE_KEY + '_time');
    var timeAgo = cacheTime ? timeSince(new Date(cacheTime)) : 'recently';
    setStatus('', 'Offline');
    showOfflineBanner(timeAgo);
    renderSongs(); updateLanguages();
  } else { setStatus('', 'Offline'); showOfflineBanner(null); }
}
function showOfflineBanner(timeAgo) {
  if(offlineBannerShown) return;
  offlineBannerShown = true;
  var banner = document.createElement('div');
  banner.id = 'offlineBanner';
  banner.style.cssText = 'background:rgba(196,151,42,0.1);border:1px solid rgba(196,151,42,0.3);border-radius:10px;padding:12px 18px;margin-bottom:16px;font-size:13px;color:#E8C36A;display:flex;align-items:center;gap:10px;';
  banner.innerHTML = '📴 <span>You\'re offline. ' + (timeAgo ? 'Showing songs cached ' + timeAgo + '.' : 'No cached songs — connect to internet first.') + '</span>';
  var container = document.querySelector('.container');
  if(container) container.insertBefore(banner, container.firstChild);
}
function hideOfflineBanner() {
  var banner = document.getElementById('offlineBanner');
  if(banner) { banner.remove(); offlineBannerShown = false; }
}
function timeSince(date) {
  var s = Math.floor((new Date() - date) / 1000);
  if(s < 60) return 'just now';
  var m = Math.floor(s/60); if(m < 60) return m + ' min ago';
  var h = Math.floor(m/60); if(h < 24) return h + ' hr ago';
  return Math.floor(h/24) + ' day(s) ago';
}
window.addEventListener('offline', function(){ setStatus('', 'Offline'); loadFromCache(); });
window.addEventListener('online', function(){
  setStatus('live', 'Live');
  hideOfflineBanner();
  if(firebaseReady && db) startRealtimeListener();
});

// ─── Excel Import ──────────────────────────
async function importFromExcel(input) {
  var file = input.files[0];
  if(!file) return;

  var useAIClean = document.getElementById('cleanWithAI').checked;

  var progressBox = document.getElementById('excelImportProgress');
  var statusEl = document.getElementById('excelImportStatus');
  var barEl = document.getElementById('excelProgressBar');
  if(progressBox) progressBox.style.display = 'block';
  if(statusEl) statusEl.textContent = 'Reading Excel file...';
  if(barEl) barEl.style.width = '10%';

  try {
    var arrayBuffer = await file.arrayBuffer();
    var workbook = XLSX.read(arrayBuffer, { type: 'array' });
    var sheetName = workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    var headerRowIdx = -1;
    for(var i = 0; i < rows.length; i++) {
      var firstCell = String(rows[i][0] || '').toLowerCase();
      if(firstCell.includes('title')) { headerRowIdx = i; break; }
    }
    if(headerRowIdx < 0) { showToast('Could not find header row in Excel', 'error'); if(progressBox) progressBox.style.display='none'; return; }

    var headers = rows[headerRowIdx].map(function(h){ return String(h).toLowerCase().trim(); });
    var dataRows = rows.slice(headerRowIdx + 2);
    var isNonEnglish = headers.some(function(h){ return h.includes('original') || h.includes('transliteration'); });

    function colIdx(keywords) {
      for(var k = 0; k < keywords.length; k++)
        for(var j = 0; j < headers.length; j++)
          if(headers[j].includes(keywords[k])) return j;
      return -1;
    }

    var iTitle  = colIdx(['title']);
    var iArtist = colIdx(['artist']);
    var iLang   = colIdx(['language']);
    var iKey    = colIdx(['key', 'tempo']);
    var iYt     = colIdx(['youtube', 'link']);
    var iLyrics = isNonEnglish ? -1 : colIdx(['lyrics']);
    var iOrig   = isNonEnglish ? colIdx(['original']) : -1;
    var iTrans  = isNonEnglish ? colIdx(['transliteration', 'phonetic', 'english trans']) : -1;

    var songsToAdd = [];
    dataRows.forEach(function(row) {
      var title = String(row[iTitle] || '').trim();
      if(!title) return;
      var songData = {
        title: title,
        artist: String(row[iArtist] || '').trim(),
        language: String(row[iLang] || (isNonEnglish ? '' : 'English')).trim() || 'English',
        key: iKey >= 0 ? String(row[iKey] || '').trim() : '',
        youtube: iYt >= 0 ? String(row[iYt] || '').trim() : '',
        dateAdded: new Date().toISOString()
      };
      if(isNonEnglish) {
        songData.lyricsOriginal = iOrig >= 0 ? String(row[iOrig] || '').trim() : '';
        songData.lyricsTransliteration = iTrans >= 0 ? String(row[iTrans] || '').trim() : '';
        if(!songData.lyricsOriginal && !songData.lyricsTransliteration) return;
      } else {
        songData.lyricsEnglish = iLyrics >= 0 ? String(row[iLyrics] || '').trim() : '';
        if(!songData.lyricsEnglish) return;
      }
      songsToAdd.push(songData);
    });

    if(songsToAdd.length === 0) { showToast('No valid songs found', 'error'); if(progressBox) progressBox.style.display='none'; input.value=''; return; }

    if(statusEl) statusEl.textContent = 'Found ' + songsToAdd.length + ' songs. Checking for duplicates...';
    if(barEl) barEl.style.width = '20%';

    var existingKeys = songs.map(function(s){ return (s.title + '|' + s.artist).toLowerCase(); });
    var newSongs = songsToAdd.filter(function(s){
      return !existingKeys.includes((s.title + '|' + s.artist).toLowerCase());
    });
    var skipped = songsToAdd.length - newSongs.length;

    if(newSongs.length === 0) { showToast('All ' + songsToAdd.length + ' songs already exist', 'error'); if(progressBox) progressBox.style.display='none'; input.value=''; return; }

    if(useAIClean) {
      if(statusEl) statusEl.textContent = '✨ AI cleaning ' + newSongs.length + ' songs...';
      if(barEl) barEl.style.width = '25%';
      for(var k = 0; k < newSongs.length; k++) {
        newSongs[k] = await cleanSongWithAI(newSongs[k]);
        var cleanPct = 25 + Math.round(((k+1) / newSongs.length) * 25);
        if(barEl) barEl.style.width = cleanPct + '%';
        if(statusEl) statusEl.textContent = '✨ AI cleaning... ' + (k+1) + ' / ' + newSongs.length;
      }
    }

    if(statusEl) statusEl.textContent = 'Importing ' + newSongs.length + ' songs' + (skipped > 0 ? ' (' + skipped + ' duplicates skipped)' : '') + '...';
    if(barEl) barEl.style.width = '55%';

    var added = 0;
    if(firebaseReady && db) {
      var batchSize = 400;
      for(var b = 0; b < newSongs.length; b += batchSize) {
        var batch = db.batch();
        var chunk = newSongs.slice(b, b + batchSize);
        chunk.forEach(function(s){ batch.set(db.collection('songs').doc(), s); });
        await batch.commit();
        added += chunk.length;
        var pct = 55 + Math.round((added / newSongs.length) * 40);
        if(barEl) barEl.style.width = pct + '%';
        if(statusEl) statusEl.textContent = 'Saving... ' + added + ' / ' + newSongs.length;
      }
    } else {
      newSongs.forEach(function(s){ songs.push(Object.assign({ id: Date.now().toString() + Math.random() }, s)); added++; });
      saveToLocalStorage(); renderSongs(); updateLanguages(); updateStats(); renderAdminList();
    }

    if(barEl) barEl.style.width = '100%';
    if(statusEl) { statusEl.textContent = '✓ Done! ' + added + ' songs imported' + (skipped > 0 ? ', ' + skipped + ' duplicates skipped.' : '.'); statusEl.style.color = '#5cd48a'; }
    showToast('Imported ' + added + ' songs!', 'success');
    setTimeout(function(){ if(progressBox) progressBox.style.display='none'; if(statusEl) statusEl.style.color=''; if(barEl) barEl.style.width='0%'; }, 4000);

  } catch(err) {
    console.error('Excel import error:', err);
    showToast('Import failed: ' + err.message, 'error');
    if(progressBox) progressBox.style.display = 'none';
  }
  input.value = '';
}
</script>
</body>
</html>
