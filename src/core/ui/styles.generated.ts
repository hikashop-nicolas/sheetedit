// GENERATED FROM src/sheetedit.css - DO NOT EDIT. Run `npm run css` after changing the CSS.
export const SHEETEDIT_CSS = `:root {
--sheetedit-bg: #ffffff;
--sheetedit-chrome: #f7f7f9;
--sheetedit-chrome2: #f1f1f4;
--sheetedit-border: #e0e2e7;
--sheetedit-text: #1c1e21;
--sheetedit-muted: #5b6270;
--sheetedit-btn: #eceef2;
--sheetedit-btn-hover: #e2e5ea;
--sheetedit-btn-border: #d3d7de;
--sheetedit-accent: #4f46e5;
--sheetedit-accent-fg: #ffffff;
--sheetedit-accent-soft: #dfe3ff;
--sheetedit-scrim: rgba(0, 0, 0, .35);
--sheetedit-shadow: rgba(0, 0, 0, .2);
--sheetedit-input-bg: #ffffff;
--sheetedit-input-fg: #1c1e21;
--sheetedit-grid-bg: #e9e9ec;
--sheetedit-cell-bg: #ffffff;
--sheetedit-cell-fg: #1a1a1a;
--sheetedit-gridline: #e3e3e6;
--sheetedit-head-bg: #f1f1f4;
--sheetedit-head-fg: #555555;
--sheetedit-head-border: #d4d4d8;
--sheetedit-head-hover: #e3e3e8;
--sheetedit-cell-focus-bg: #eef0ff;
--sheetedit-zebra: #f6f6f8;
--sheetedit-link: #2563eb;
--sheetedit-flat-bg: #eef0f4;
--sheetedit-flat-hover: #e3e6ec;
--sheetedit-flat-fg: #555555;
--sheetedit-flat-fg-hover: #111111;
--sheetedit-danger: #d33d3d;
--sheetedit-danger-strong: #e03131;
--sheetedit-invalid: #e0533d;
--sheetedit-comment-mark: #d9534f;
--sheetedit-ok: #3fb950;
--sheetedit-neutral: #c0c4cc;
--sheetedit-faint: #8a8f98;
--sheetedit-error-bg: #fde8e8;
--sheetedit-error-fg: #8a2222;
--sheetedit-error-text: #c0392b;
--sheetedit-notice-bg: #fdf1d8;
--sheetedit-notice-fg: #6b4a10;
--sheetedit-notice-border: #e6cf9a;
--sheetedit-pagebreak: #4f7bd6;
--sheetedit-printarea: #2f6f3f;
--sheetedit-paper: #ffffff;
--sheetedit-ink: #000000;
--sheetedit-print-rule: #b0b0b0;
--sheetedit-print-head-bg: #efefef;
--sheetedit-overlay-bg: #ffffff;
--sheetedit-overlay-border: #d0d0d6;
--sheetedit-handle-border: #ffffff;
}
@media (prefers-color-scheme: dark) {
:root:not([data-theme="light"]) {
--sheetedit-bg: #1f2227;
--sheetedit-chrome: #2b2f36;
--sheetedit-chrome2: #23262c;
--sheetedit-border: #1c1f24;
--sheetedit-text: #e6e6e6;
--sheetedit-muted: #cfd3da;
--sheetedit-btn: #3a3f47;
--sheetedit-btn-hover: #454b54;
--sheetedit-btn-border: #4a4f57;
--sheetedit-accent: #6e7bff;
--sheetedit-accent-fg: #ffffff;
--sheetedit-accent-soft: #2f3560;
--sheetedit-scrim: rgba(0, 0, 0, .5);
--sheetedit-shadow: rgba(0, 0, 0, .55);
--sheetedit-input-bg: #1c1f24;
--sheetedit-input-fg: #e7eaf0;
--sheetedit-grid-bg: #16181c;
--sheetedit-cell-bg: #1e2126;
--sheetedit-cell-fg: #e3e6ea;
--sheetedit-gridline: #2f343b;
--sheetedit-head-bg: #262a31;
--sheetedit-head-fg: #b9c0cb;
--sheetedit-head-border: #343a43;
--sheetedit-head-hover: #2f353d;
--sheetedit-cell-focus-bg: #262c4a;
--sheetedit-zebra: #22252b;
--sheetedit-link: #7aa2ff;
--sheetedit-flat-bg: #2a2f36;
--sheetedit-flat-hover: #343a43;
--sheetedit-flat-fg: #b9c0cb;
--sheetedit-flat-fg-hover: #e6e6e6;
--sheetedit-neutral: #4a5059;
--sheetedit-faint: #8a8f98;
--sheetedit-error-bg: #7a2b2b;
--sheetedit-error-fg: #ffd7d7;
--sheetedit-error-text: #ff8a8a;
--sheetedit-notice-bg: #4a3410;
--sheetedit-notice-fg: #ffe2b0;
--sheetedit-notice-border: #6d4c16;
--sheetedit-pagebreak: #7aa2ff;
--sheetedit-printarea: #57a46b;
--sheetedit-overlay-bg: #22262c;
--sheetedit-overlay-border: #3a4049;
--sheetedit-handle-border: #e6e6e6;
}
}
:root[data-theme="dark"] {
--sheetedit-bg: #1f2227;
--sheetedit-chrome: #2b2f36;
--sheetedit-chrome2: #23262c;
--sheetedit-border: #1c1f24;
--sheetedit-text: #e6e6e6;
--sheetedit-muted: #cfd3da;
--sheetedit-btn: #3a3f47;
--sheetedit-btn-hover: #454b54;
--sheetedit-btn-border: #4a4f57;
--sheetedit-accent: #6e7bff;
--sheetedit-accent-fg: #ffffff;
--sheetedit-accent-soft: #2f3560;
--sheetedit-scrim: rgba(0, 0, 0, .5);
--sheetedit-shadow: rgba(0, 0, 0, .55);
--sheetedit-input-bg: #1c1f24;
--sheetedit-input-fg: #e7eaf0;
--sheetedit-grid-bg: #16181c;
--sheetedit-cell-bg: #1e2126;
--sheetedit-cell-fg: #e3e6ea;
--sheetedit-gridline: #2f343b;
--sheetedit-head-bg: #262a31;
--sheetedit-head-fg: #b9c0cb;
--sheetedit-head-border: #343a43;
--sheetedit-head-hover: #2f353d;
--sheetedit-cell-focus-bg: #262c4a;
--sheetedit-zebra: #22252b;
--sheetedit-link: #7aa2ff;
--sheetedit-flat-bg: #2a2f36;
--sheetedit-flat-hover: #343a43;
--sheetedit-flat-fg: #b9c0cb;
--sheetedit-flat-fg-hover: #e6e6e6;
--sheetedit-neutral: #4a5059;
--sheetedit-faint: #8a8f98;
--sheetedit-error-bg: #7a2b2b;
--sheetedit-error-fg: #ffd7d7;
--sheetedit-error-text: #ff8a8a;
--sheetedit-notice-bg: #4a3410;
--sheetedit-notice-fg: #ffe2b0;
--sheetedit-notice-border: #6d4c16;
--sheetedit-pagebreak: #7aa2ff;
--sheetedit-printarea: #57a46b;
--sheetedit-overlay-bg: #22262c;
--sheetedit-overlay-border: #3a4049;
--sheetedit-handle-border: #e6e6e6;
}
.sheetedit-wrap { position:relative; display:flex; flex-direction:column; height:100%; background:var(--sheetedit-bg, #1f2227); color:var(--sheetedit-text, #e6e6e6); font:13px system-ui, sans-serif; }
.sheetedit-toolbar { display:flex; flex-wrap:nowrap; overflow:hidden; align-items:center; gap:5px; padding:5px 8px; background:var(--sheetedit-chrome, #2b2f36); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.sheetedit-btn {
font:inherit; font-size:13px; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
border-radius:6px; padding:4px 9px; cursor:pointer; min-width:32px; line-height:1.1;
}
.sheetedit-btn:hover { background:var(--sheetedit-btn-hover, #454b54); }
.sheetedit-btn.is-active { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); }
.sheetedit-btn.is-active:hover { background:var(--sheetedit-accent, #6e7bff); }
.sheetedit-tb-moremenu { flex-direction:column; align-items:stretch; gap:2px; max-height:70vh; overflow-y:auto; }
.sheetedit-more-item { display:flex; align-items:center; gap:9px; justify-content:flex-start; text-align:left; min-width:170px; }
.sheetedit-btn:focus-visible { outline:2px solid var(--sheetedit-accent, #6e7bff); outline-offset:1px; }
.sheetedit-tb-sep { width:1px; align-self:stretch; background:var(--sheetedit-btn-border, #4a4f57); margin:1px 3px; }
.sheetedit-color { width:30px; height:28px; padding:0; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px; background:var(--sheetedit-btn, #3a3f47); cursor:pointer; }
.sheetedit-tb-select {
font:inherit; font-size:13px; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
border-radius:6px; padding:4px 4px; cursor:pointer; max-width:64px; height:28px;
}
.sheetedit-tb-select:focus-visible { outline:2px solid var(--sheetedit-accent, #6e7bff); outline-offset:1px; }
.sheetedit-btn svg { display:block; width:16px; height:16px; }
.sheetedit-table th.colhead, .sheetedit-table th.rownum, .sheetedit-table th.corner { cursor:pointer; }
.sheetedit-table th.colhead:hover, .sheetedit-table th.rownum:hover, .sheetedit-table th.corner:hover { background:var(--sheetedit-head-hover); }
.sheetedit-table td.sheetedit-sel input { background:rgba(110,123,255,0.18); }
.sheetedit-table td.has-ruby { position:relative; }
.sheetedit-table td.has-ruby:not(:focus-within) input { color:transparent !important; }
.sheetedit-table td.has-ruby .sheetedit-ruby {
position:absolute; inset:0; display:flex; align-items:center; padding:0 6px;
pointer-events:none; overflow:hidden; white-space:nowrap; line-height:1.05;
}
.sheetedit-table td.has-ruby .sheetedit-ruby rt { font-size:0.6em; line-height:1; user-select:none; }
.sheetedit-table td.has-ruby:focus-within .sheetedit-ruby { display:none; }
.sheetedit-table td.has-wrap { position:relative; }
.sheetedit-table td.has-wrap:not(:focus-within) input { color:transparent !important; }
.sheetedit-table td.has-wrap .sheetedit-cellwrap {
position:absolute; inset:0; padding:3px 8px; white-space:pre-wrap; word-break:break-word;
overflow:hidden; pointer-events:none; line-height:1.3; color:var(--sheetedit-cell-fg);
}
.sheetedit-table td.has-wrap:focus-within .sheetedit-cellwrap { display:none; }
.sheetedit-table td.has-rich { position:relative; }
.sheetedit-table td.has-rich:not(:focus-within) input { color:transparent !important; }
.sheetedit-table td.has-rich .sheetedit-cellrich {
position:absolute; inset:0; padding:1px 8px; display:flex; align-items:center; white-space:pre;
overflow:hidden; pointer-events:none; color:var(--sheetedit-cell-fg);
}
.sheetedit-table td.has-rich:focus-within .sheetedit-cellrich { display:none; }
.sheetedit-table td.has-link { position:relative; }
.sheetedit-table td.has-link input:not(:focus) { color:var(--sheetedit-link); text-decoration:underline; }
.sheetedit-linkbtn { position:absolute; top:1px; right:1px; z-index:3; display:inline-flex; align-items:center; justify-content:center; width:15px; height:15px; padding:0; border:0; border-radius:3px; background:transparent; color:var(--sheetedit-link); cursor:pointer; opacity:.8; }
.sheetedit-linkbtn:hover { opacity:1; background:rgba(37,99,235,0.14); }
.sheetedit-table td.has-dv { position:relative; }
.sheetedit-dvbtn { position:absolute; top:0; right:0; bottom:0; z-index:3; visibility:hidden; display:inline-flex; align-items:center; justify-content:center; width:17px; padding:0; border:0; border-left:1px solid var(--sheetedit-head-border); background:var(--sheetedit-flat-bg); color:var(--sheetedit-head-fg); cursor:pointer; }
.sheetedit-dvbtn:hover { background:var(--sheetedit-flat-hover); color:var(--sheetedit-flat-fg-hover); }
.sheetedit-table td.has-dv:hover .sheetedit-dvbtn, .sheetedit-table td.has-dv:focus-within .sheetedit-dvbtn, .sheetedit-table td.has-dv.sheetedit-sel .sheetedit-dvbtn { visibility:visible; }
.sheetedit-table td.sheetedit-dv-invalid { box-shadow: inset 0 0 0 2px var(--sheetedit-invalid); }
.sheetedit-dvmenu { min-width:120px; max-height:240px; overflow-y:auto; }
.sheetedit-table td.has-spark { position:relative; overflow:hidden; }
.sheetedit-spark { position:absolute; left:2px; top:2px; width:calc(100% - 4px); height:calc(100% - 4px); z-index:0; pointer-events:none; }
.sheetedit-table td.has-spark input { position:relative; z-index:1; background:transparent; }
.sheetedit-filterbtn { position:absolute; top:0; right:0; bottom:0; z-index:3; visibility:hidden; display:inline-flex; align-items:center; justify-content:center; width:17px; padding:0; border:0; border-left:1px solid var(--sheetedit-head-border); background:var(--sheetedit-flat-bg); color:var(--sheetedit-head-fg); cursor:pointer; }
.sheetedit-filterbtn:hover { background:var(--sheetedit-flat-hover); color:var(--sheetedit-flat-fg-hover); }
.sheetedit-table td.has-filter { position:relative; }
.sheetedit-table td.has-filter .sheetedit-filterbtn { visibility:visible; }
.sheetedit-table td.sheetedit-filter-on .sheetedit-filterbtn { color:var(--sheetedit-accent, #6e7bff); background:var(--sheetedit-accent-soft); }
.sheetedit-filtermenu { min-width:190px; }
.sheetedit-pop-sep { height:1px; margin:4px 6px; background:var(--sheetedit-border, #3a3f47); }
.sheetedit-filter-list { max-height:220px; overflow-y:auto; padding:2px 0; }
.sheetedit-filter-opt { display:flex; align-items:center; gap:7px; padding:4px 11px; font-size:13px; cursor:pointer; }
.sheetedit-filter-opt:hover { background:var(--sheetedit-btn, #3a3f47); }
.sheetedit-filter-foot { display:flex; justify-content:flex-end; padding:6px 8px 4px; border-top:1px solid var(--sheetedit-border, #3a3f47); }
.sheetedit-table td.has-cfbar { position:relative; }
.sheetedit-cfbar { position:absolute; left:1px; top:2px; bottom:2px; z-index:0; border-radius:1px; opacity:.85; pointer-events:none; }
.sheetedit-table td.has-cfbar input { position:relative; z-index:1; background:transparent; }
.sheetedit-table td.has-cficon { position:relative; }
.sheetedit-cficon { position:absolute; left:3px; top:50%; transform:translateY(-50%); z-index:1; display:inline-flex; pointer-events:none; }
.sheetedit-table td.has-cficon input { padding-left:18px; }
.sheetedit-table td.has-comment { position:relative; }
.sheetedit-commark { position:absolute; top:0; right:0; z-index:2; width:0; height:0; border-top:6px solid var(--sheetedit-comment-mark); border-left:6px solid transparent; pointer-events:none; }
.sheetedit-compop { position:fixed; z-index:40; max-width:260px; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e7eaf0); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 8px 24px rgba(0,0,0,.45); padding:8px 10px; font:12px/1.45 system-ui,sans-serif; }
.sheetedit-comitem + .sheetedit-comitem { margin-top:7px; padding-top:7px; border-top:1px solid var(--sheetedit-border, #1c1f24); }
.sheetedit-comauthor { font-weight:600; margin-bottom:2px; }
.sheetedit-comtext { white-space:pre-wrap; color:var(--sheetedit-muted, #cfd3da); }
.sheetedit-chartlayer { position:absolute; overflow:hidden; pointer-events:none; z-index:6; }
.sheetedit-chartlayer-inner { position:absolute; top:0; left:0; }
.sheetedit-chartbox { position:absolute; pointer-events:auto; background:var(--sheetedit-cell-bg); border:1px solid var(--sheetedit-overlay-border); border-radius:3px; box-shadow:0 1px 5px rgba(0,0,0,.15); padding:5px; box-sizing:border-box; }
.sheetedit-chartbox.sel { border-color:var(--sheetedit-accent, #6e7bff); box-shadow:0 0 0 2px var(--sheetedit-accent, #6e7bff); }
.sheetedit-furi-pop { min-width:180px; gap:6px; }
.sheetedit-furi-input { font:inherit; font-size:13px; padding:6px 8px; border-radius:5px; border:1px solid var(--sheetedit-btn-border,#4a4f57); background:var(--sheetedit-btn,#3a3f47); color:var(--sheetedit-text,#e6e6e6); }
.sheetedit-furi-row { display:flex; gap:4px; }
.sheetedit-furi-row .sheetedit-pop-item { flex:1; text-align:center; }
.sheetedit-pop { position:fixed; z-index:30; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:8px; padding:4px; box-shadow:0 6px 18px rgba(0,0,0,0.45); display:flex; flex-direction:column; min-width:130px; }
.sheetedit-pop-item { font:inherit; font-size:13px; text-align:left; background:transparent; color:var(--sheetedit-text, #e6e6e6); border:0; border-radius:5px; padding:7px 11px; cursor:pointer; }
.sheetedit-pop-item:hover { background:var(--sheetedit-btn, #3a3f47); }
.sheetedit-gridarea { flex:1; min-height:0; display:flex; flex-direction:row; }
.sheetedit-colband { min-width:0; display:flex; flex-direction:column; }
.sheetedit-colband:first-child { flex:1; }
.sheetedit-colband-right { flex:1; display:none; }
.sheetedit-grid { flex:1; min-height:0; overflow:auto; background:var(--sheetedit-grid-bg); }
.sheetedit-grid-split { border-top:1px solid var(--sheetedit-border, #c8ccd2); }
.sheetedit-grid-right { border-left:1px solid var(--sheetedit-border, #c8ccd2); }
.sheetedit-grid:focus { outline:none; }
table.sheetedit-table { border-collapse:collapse; table-layout:fixed; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; }
.sheetedit-table th, .sheetedit-table td { padding:0; margin:0; }
.sheetedit-table th { border:1px solid var(--sheetedit-head-border); }
.sheetedit-table td { background:var(--sheetedit-cell-bg); box-shadow: inset -1px -1px 0 0 var(--sheetedit-gridline); }
.sheetedit-table th {
position:sticky; top:0; z-index:8; background:var(--sheetedit-head-bg); color:var(--sheetedit-head-fg); font-weight:600;
padding:3px 8px; text-align:center; user-select:none;
}
.sheetedit-table th.corner { left:0; z-index:9; }
.sheetedit-table th.rownum { position:sticky; left:0; z-index:5; top:auto; text-align:right; background:var(--sheetedit-head-bg); }
.sheetedit-table td.frz { background:var(--sheetedit-cell-bg); }
.sheetedit-table td.va-top { vertical-align:top; }
.sheetedit-table td.va-bottom { vertical-align:bottom; }
.sheetedit-table td.va-top > input, .sheetedit-table td.va-bottom > input { display:inline-block; }
.sheetedit-colgrip { position:absolute; top:0; right:-4px; width:9px; height:100%; cursor:col-resize; z-index:4; touch-action:none; }
.sheetedit-rowgrip { position:absolute; left:0; bottom:-4px; width:100%; height:9px; cursor:row-resize; z-index:4; touch-action:none; }
.sheetedit-colgrip:hover { box-shadow:inset -2px 0 0 0 var(--sheetedit-accent, #6e7bff); }
.sheetedit-rowgrip:hover { box-shadow:inset 0 -2px 0 0 var(--sheetedit-accent, #6e7bff); }
.sheetedit-table input {
border:0; background:transparent; color:var(--sheetedit-cell-fg); font:inherit; padding:3px 8px;
width:100%; box-sizing:border-box; outline:none;
}
.sheetedit-table td.num input { text-align:right; font-variant-numeric:tabular-nums; }
.sheetedit-table td.sheetedit-fillsrc { position:relative; }
.sheetedit-fillhandle {
position:absolute; right:-4px; bottom:-4px; width:8px; height:8px; z-index:5;
background:var(--sheetedit-accent, #6e7bff); border:1px solid var(--sheetedit-handle-border); cursor:crosshair; touch-action:none;
}
.sheetedit-table td.sheetedit-fillprev input { background:rgba(110,123,255,0.10); }
.sheetedit-table td.sheetedit-fillprev { box-shadow: inset 0 0 0 1px var(--sheetedit-accent, #6e7bff); }
.sheetedit-findbar { display:flex; align-items:center; gap:6px; padding:5px 8px; background:var(--sheetedit-chrome2, #23262c); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.sheetedit-findbar[hidden] { display:none; }
.sheetedit-findbar input { flex:0 1 180px; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); font:13px system-ui,sans-serif; padding:4px 8px; }
.sheetedit-findcount { color:var(--sheetedit-muted, #aab2bf); font:12px system-ui,sans-serif; min-width:64px; }
.sheetedit-table td.sheetedit-calcerr { position:relative; }
.sheetedit-table td.sheetedit-calcerr::after {
content:""; position:absolute; top:0; right:0; z-index:1; pointer-events:none;
border:5px solid transparent; border-top-color:var(--sheetedit-danger); border-right-color:var(--sheetedit-danger);
}
.sheetedit-table input:focus { box-shadow:inset 0 0 0 2px var(--sheetedit-accent, #6e7bff); background:var(--sheetedit-cell-focus-bg); }
.sheetedit-tb-slot { display:inline-flex; align-items:center; gap:5px; }
.sheetedit-tb-groupmenu { position:absolute; z-index:30; display:flex; align-items:center; gap:5px; padding:6px 8px; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
.sheetedit-fxbar { display:flex; align-items:center; gap:6px; padding:4px 8px; background:var(--sheetedit-chrome2, #23262c); border-bottom:1px solid var(--sheetedit-border, #1c1f24); position:relative; }
.sheetedit-fxref { min-width:52px; text-align:center; font:12px/1.6 ui-monospace,monospace; color:var(--sheetedit-muted, #aab2bf); background:var(--sheetedit-chrome, #2b2f36); border-radius:5px; padding:2px 6px; }
.sheetedit-fxbtns { position:relative; display:inline-flex; gap:2px; }
.sheetedit-fxsum { font-weight:700; }
.sheetedit-fxmenu { position:absolute; top:100%; left:0; z-index:30; display:flex; flex-direction:column; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); padding:4px; }
.sheetedit-fxmenu-item { background:none; border:0; color:var(--sheetedit-text, #e7eaf0); text-align:left; padding:5px 12px; font:12px system-ui,sans-serif; cursor:pointer; border-radius:5px; }
.sheetedit-fxmenu-item:hover { background:var(--sheetedit-btn, #3a4047); }
.sheetedit-fxinput { flex:1; min-width:60px; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); font:13px ui-monospace,monospace; padding:4px 8px; }
.sheetedit-fxbar.is-picking .sheetedit-fxinput { border-color:var(--sheetedit-accent, #4f8ef7); }
.sheetedit-fxmenu[hidden], .sheetedit-tb-groupmenu[hidden] { display:none; }
.sheetedit-fxassist { display:inline-flex; align-items:center; justify-content:center; }
.sheetedit-fxa-pop { position:absolute; z-index:40; width:min(380px,92%); box-sizing:border-box; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e7eaf0); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 10px 34px rgba(0,0,0,.5); padding:12px; display:flex; flex-direction:column; gap:8px; font:13px/1.4 system-ui,sans-serif; }
.sheetedit-fxa-pop[hidden] { display:none; }
.sheetedit-fxa-title { font-weight:600; font-size:14px; }
.sheetedit-fxa-desc { width:100%; box-sizing:border-box; resize:vertical; padding:7px; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:6px; background:var(--sheetedit-chrome2, #23262c); color:var(--sheetedit-text, #e7eaf0); font:inherit; }
.sheetedit-fxa-progress { color:var(--sheetedit-muted, #aab2bf); font-size:12px; min-height:15px; }
.sheetedit-fxa-rlabel { display:flex; flex-direction:column; gap:4px; color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
.sheetedit-fxa-result { width:100%; box-sizing:border-box; padding:6px 8px; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:6px; background:var(--sheetedit-border, #1c1f24); color:var(--sheetedit-text, #e7eaf0); font:13px ui-monospace,monospace; }
.sheetedit-fxa-actions { display:flex; gap:8px; justify-content:flex-end; }
.sheetedit-fxa-btn { padding:6px 14px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:7px; background:var(--sheetedit-chrome2, #23262c); color:var(--sheetedit-text, #e7eaf0); font:inherit; cursor:pointer; }
.sheetedit-fxa-btn:hover:not(:disabled) { border-color:var(--sheetedit-accent, #6e7bff); }
.sheetedit-fxa-btn.is-primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); }
.sheetedit-fxa-btn:disabled { opacity:.45; cursor:default; }
.sheetedit-qp-pop { position:absolute; z-index:40; width:min(460px,94%); box-sizing:border-box; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e7eaf0); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 10px 34px rgba(0,0,0,.5); padding:12px; display:flex; flex-direction:column; gap:8px; font:13px/1.4 system-ui,sans-serif; max-height:70vh; }
.sheetedit-qp-pop[hidden] { display:none; }
.sheetedit-qp-title { font-weight:600; font-size:14px; }
.sheetedit-qp-body { display:flex; flex-direction:column; gap:8px; overflow:auto; min-height:0; }
.sheetedit-qp-row { border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; padding:8px; display:flex; flex-direction:column; gap:5px; }
.sheetedit-qp-rowhead { display:flex; align-items:center; gap:8px; }
.sheetedit-qp-name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:600; }
.sheetedit-qp-btn { padding:4px 11px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px; background:var(--sheetedit-chrome2, #23262c); color:var(--sheetedit-text, #e7eaf0); font:inherit; cursor:pointer; }
.sheetedit-qp-btn:hover:not(:disabled) { border-color:var(--sheetedit-accent, #6e7bff); }
.sheetedit-qp-btn:disabled { opacity:.45; cursor:default; }
.sheetedit-qp-status { color:var(--sheetedit-muted, #aab2bf); font-size:12px; min-height:14px; }
.sheetedit-qp-m { margin:0; max-height:180px; overflow:auto; padding:7px 9px; border-radius:6px; background:var(--sheetedit-border, #1c1f24); color:var(--sheetedit-text, #e7eaf0); font:12px/1.5 ui-monospace,monospace; white-space:pre-wrap; }
.sheetedit-qp-note { color:var(--sheetedit-muted, #aab2bf); font-size:11px; }
.sheetedit-qp-attach { display:inline-block; margin:4px 0; font-size:11px; color:var(--sheetedit-accent, #6cf); cursor:pointer; }
.sheetedit-qp-medit { display:flex; flex-direction:column; gap:4px; margin-top:4px; }
.sheetedit-qp-medit button { align-self:flex-start; }
.sheetedit-qp-mwrap {
--se-code-bg:#ffffff; --se-code-border:#c9ccd4; --se-code-fg:#24292e; --se-code-caret:#111;
--se-kw:#0000c8; --se-fn:#795e26; --se-str:#a31515; --se-num:#098658; --se-com:#2e8b57; --se-op:#555; --se-id:#001080;
position:relative; height:170px; border:1px solid var(--se-code-border); border-radius:4px; overflow:hidden; background:var(--se-code-bg);
}
.sheetedit-qp-mhl, .sheetedit-qp-medit textarea.sheetedit-qp-m {
position:absolute; inset:0; margin:0; box-sizing:border-box; padding:6px; border:0;
font-family:ui-monospace,Menlo,Consolas,monospace; font-size:11px; line-height:1.5;
white-space:pre; overflow:auto; tab-size:4; -moz-tab-size:4;
}
.sheetedit-qp-mhl { color:var(--se-code-fg); pointer-events:none; z-index:0; }
.sheetedit-qp-mhl code { font:inherit; white-space:inherit; }
.sheetedit-qp-medit textarea.sheetedit-qp-m { z-index:1; background:transparent; color:transparent; caret-color:var(--se-code-caret); resize:none; outline:none; }
.sheetedit-qp-mhl .mtok-kw { color:var(--se-kw); font-weight:600; }
.sheetedit-qp-mhl .mtok-fn { color:var(--se-fn); }
.sheetedit-qp-mhl .mtok-str { color:var(--se-str); }
.sheetedit-qp-mhl .mtok-num { color:var(--se-num); }
.sheetedit-qp-mhl .mtok-com { color:var(--se-com); font-style:italic; }
.sheetedit-qp-mhl .mtok-op { color:var(--se-op); }
.sheetedit-qp-mhl .mtok-id { color:var(--se-id); }
@media (prefers-color-scheme: dark) {
:root:not([data-theme="light"]) .sheetedit-qp-mwrap {
--se-code-bg:#1e2228; --se-code-border:#3a3f4b; --se-code-fg:#d6dae0; --se-code-caret:#e6e6e6;
--se-kw:#7aa2ff; --se-fn:#d7ba7d; --se-str:#ce9178; --se-num:#b5cea8; --se-com:#7fb37f; --se-op:#b0b6c0; --se-id:#9cdcfe;
}
}
:root[data-theme="dark"] .sheetedit-qp-mwrap {
--se-code-bg:#1e2228; --se-code-border:#3a3f4b; --se-code-fg:#d6dae0; --se-code-caret:#e6e6e6;
--se-kw:#7aa2ff; --se-fn:#d7ba7d; --se-str:#ce9178; --se-num:#b5cea8; --se-com:#7fb37f; --se-op:#b0b6c0; --se-id:#9cdcfe;
}
.sheetedit-fmtmenu { flex-direction:column; align-items:stretch; gap:2px; }
.sheetedit-fmtmenu .sheetedit-btn { text-align:left; justify-content:flex-start; }
.sheetedit-floatbar { position:fixed; z-index:40; display:flex; align-items:center; gap:2px; padding:4px 6px; background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
.sheetedit-floatbar[hidden] { display:none; }
.sheetedit-floatbar-sep { width:1px; align-self:stretch; margin:2px 3px; background:var(--sheetedit-border, #1c1f24); }
.sheetedit-error { background:var(--sheetedit-error-bg); color:var(--sheetedit-error-fg); padding:10px 14px; font:13px/1.5 system-ui,sans-serif; }
.sheetedit-notice {
position:absolute; left:50%; bottom:44px; transform:translateX(-50%); z-index:60; max-width:min(90%,420px);
padding:7px 13px; border-radius:7px; background:var(--sheetedit-notice-bg); color:var(--sheetedit-notice-fg); border:1px solid var(--sheetedit-notice-border);
font:12.5px/1.45 system-ui,sans-serif; box-shadow:0 6px 18px rgba(0,0,0,.35); pointer-events:none;
}
.sheetedit-notice[hidden] { display:none; }
.sheetedit-grid td.locked > input { cursor:default; }
.sheetedit-grid td.pgbrk-top, .sheetedit-grid th.pgbrk-top { border-top:2px dashed var(--sheetedit-pagebreak) !important; }
.sheetedit-grid td.pgbrk-left, .sheetedit-grid th.pgbrk-left { border-left:2px dashed var(--sheetedit-pagebreak) !important; }
.sheetedit-grid td.pa-top { border-top:2px solid var(--sheetedit-printarea) !important; }
.sheetedit-grid td.pa-bottom { border-bottom:2px solid var(--sheetedit-printarea) !important; }
.sheetedit-grid td.pa-left { border-left:2px solid var(--sheetedit-printarea) !important; }
.sheetedit-grid td.pa-right { border-right:2px solid var(--sheetedit-printarea) !important; }
.sheetedit-grid td.pa-out { background-image:linear-gradient(rgba(120,120,130,.10), rgba(120,120,130,.10)); }
.sheetedit-tabs { display:flex; align-items:center; gap:2px; padding:5px 8px; background:var(--sheetedit-chrome, #2b2f36); border-top:1px solid var(--sheetedit-border, #1c1f24); overflow-x:auto; }
.sheetedit-tab {
font:inherit; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-muted, #cfd3da); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-bottom:none;
border-radius:5px 5px 0 0; padding:4px 12px; cursor:pointer; white-space:nowrap;
}
.sheetedit-tab[aria-selected="true"] { background:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); border-color:var(--sheetedit-accent, #6e7bff); }
.sheetedit-tab:focus-visible { outline:2px solid var(--sheetedit-accent-fg); outline-offset:1px; }
.sheetedit-tab-rename { font:inherit; width:9ch; min-width:60px; box-sizing:border-box; background:var(--sheetedit-bg, #1f2227); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-accent, #6e7bff); border-radius:4px; padding:2px 5px; }
.sheetedit-tab-add { display:inline-flex; align-items:center; justify-content:center; flex:none; width:26px; height:26px; margin-left:4px; padding:0; cursor:pointer; background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-muted, #cfd3da); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; }
.sheetedit-tab-add:hover { background:var(--sheetedit-btn-hover, #454b54); color:var(--sheetedit-text, #e6e6e6); }
.sheetedit-tabmenu { bottom:40px; }
.sheetedit-modal { position:fixed; inset:0; z-index:70; display:flex; align-items:center; justify-content:center; background:var(--sheetedit-scrim, rgba(0,0,0,.45)); }
.sheetedit-card {
width:min(420px,94%); max-height:90vh; overflow-y:auto; padding:16px; font:13px system-ui,sans-serif;
background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e6e6e6);
border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 14px 44px var(--sheetedit-shadow, rgba(0,0,0,.5));
}
.sheetedit-card.is-wide { width:min(840px,96%); max-height:88vh; overflow:auto; }
.sheetedit-card h3 { margin:0 0 12px; font-size:15px; }
.sheetedit-card.is-wide h3 { margin:0 0 4px; }
.sheetedit-note { margin:0 0 12px; color:var(--sheetedit-muted, #aab2bf); font-size:13px; }
.sheetedit-subtle { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
.sheetedit-field { display:flex; flex-direction:column; gap:4px; margin-bottom:10px; font-size:13px; }
.sheetedit-field.is-inline { flex-direction:row; align-items:center; gap:7px; }
.sheetedit-field > span { color:var(--sheetedit-muted, #aab2bf); }
.sheetedit-input {
font:inherit; padding:6px 8px; border-radius:5px;
background:var(--sheetedit-input-bg, #1c1f24); color:var(--sheetedit-input-fg, #e7eaf0); border:1px solid var(--sheetedit-btn, #3a4047);
}
.sheetedit-color { width:34px; height:26px; padding:0; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; background:none; cursor:pointer; }
.sheetedit-actions {
display:flex; justify-content:flex-end; gap:8px; margin-top:6px;
position:sticky; bottom:-16px; padding:10px 0; background:var(--sheetedit-chrome, #2b2f36);
}
.sheetedit-dlg-btn {
font:inherit; font-size:13px; padding:6px 14px; border-radius:6px; cursor:pointer;
background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
}
.sheetedit-dlg-btn.is-primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg, #fff); }
.sheetedit-field.is-hidden { display:none; }
.sheetedit-pivot-menu {
position:fixed; z-index:80; min-width:150px; padding:4px; font:13px system-ui,sans-serif;
background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e6e6e6);
border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 10px 30px var(--sheetedit-shadow, rgba(0,0,0,.5));
}
.sheetedit-pivot-menu button {
display:block; width:100%; text-align:left; padding:6px 10px; border:0; border-radius:5px;
background:none; color:inherit; font:inherit; cursor:pointer;
}
.sheetedit-pivot-menu button:disabled, .sheetedit-pivot-menu button.is-inert { cursor:default; opacity:.6; }
.sheetedit-pivot-body { display:flex; gap:16px; flex-wrap:wrap; }
.sheetedit-pivot-left { flex:1 1 360px; min-width:340px; }
.sheetedit-pivot-right { flex:1 1 300px; min-width:280px; overflow:hidden; }
.sheetedit-pivot-preview {
border:1px solid var(--sheetedit-btn, #3a4047); border-radius:6px; padding:8px; overflow:auto; max-height:260px; font-size:12px;
}
.sheetedit-pivot-preview table { border-collapse:collapse; }
.sheetedit-pivot-preview td { border:1px solid var(--sheetedit-btn, #3a4047); padding:2px 6px; white-space:nowrap; }
.sheetedit-pivot-preview td.num { text-align:right; }
.sheetedit-pivot-preview td.bold { font-weight:600; }
.sheetedit-pivot-preview td.ellipsis { border:0; padding:2px 6px; }
.sheetedit-pivot-row { display:flex; align-items:center; gap:6px; margin-bottom:8px; flex-wrap:wrap; }
.sheetedit-pivot-name { flex:1 1 84px; min-width:60px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.sheetedit-pivot-row select, .sheetedit-pivot-calc select, .sheetedit-pivot-calcitem select { max-width:128px; }
.sheetedit-pivot-calcitem select { max-width:100px; }
.sheetedit-pivot-group { margin-top:8px; }
.sheetedit-pivot-calc, .sheetedit-pivot-calcitem { display:flex; align-items:center; gap:6px; margin-bottom:6px; }
.sheetedit-pivot-calcitem { flex-wrap:wrap; }
.sheetedit-pivot-calc .name, .sheetedit-pivot-calcitem .name { flex:0 0 90px; }
.sheetedit-pivot-calcitem .name { flex-basis:80px; }
.sheetedit-pivot-calc .formula, .sheetedit-pivot-calcitem .formula { flex:1 1 auto; min-width:0; }
.sheetedit-iconbtn {
font:inherit; padding:3px 7px; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px;
background:none; color:inherit; cursor:pointer;
}
.sheetedit-addbtn {
font:inherit; font-size:12px; padding:4px 9px; border:1px dashed var(--sheetedit-btn, #3a4047); border-radius:5px;
background:none; color:var(--sheetedit-muted, #aab2bf); cursor:pointer;
}
.sheetedit-checkline { display:flex; align-items:center; gap:7px; margin-top:6px; color:var(--sheetedit-muted, #aab2bf); }
.sheetedit-actions.is-plain { position:static; margin-top:14px; padding:0; background:none; }
.sheetedit-hidden { display:none; }
.sheetedit-pivot-menu button:hover:not(.is-inert) { background:var(--sheetedit-btn, #3a3f47); }
.sheetedit-dlg-btn:disabled { opacity:.45; cursor:not-allowed; }
.sheetedit-pivot-preview.is-message { color:var(--sheetedit-muted, #aab2bf); }
.sheetedit-swatch-label { font-size:12px; color:var(--sheetedit-muted, #aab2bf); }
.sheetedit-swatch { width:26px; height:22px; padding:0; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:4px; background:none; cursor:pointer; }
.sheetedit-smallbtn {
font:inherit; font-size:13px; line-height:1; padding:2px 6px; border-radius:4px; cursor:pointer;
background:var(--sheetedit-btn, #3a3f47); color:inherit; border:1px solid var(--sheetedit-btn-border, #4a4f57);
}
.sheetedit-slicerlayer { position:absolute; overflow:hidden; pointer-events:none; z-index:6; }
.sheetedit-slicerlayer-inner { position:absolute; inset:0; }
.sheetedit-slicerbox { position:absolute; pointer-events:auto; display:flex; flex-direction:column;
background:var(--sheetedit-chrome,#fff); color:var(--sheetedit-text,#1c1f24);
border:1px solid var(--sheetedit-border,#c8ccd2); border-radius:6px;
box-shadow:0 2px 10px rgba(0,0,0,.18); font:12px system-ui,sans-serif; overflow:hidden; }
.sheetedit-slicer-head { display:flex; align-items:center; gap:6px; padding:5px 7px; font-weight:600;
border-bottom:1px solid var(--sheetedit-border,#c8ccd2); background:rgba(127,127,127,.08); }
.sheetedit-slicer-title { flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.sheetedit-slicer-clear { border:0; background:none; cursor:pointer; color:inherit; opacity:.65;
font:inherit; padding:1px 4px; border-radius:4px; }
.sheetedit-slicer-clear:hover { opacity:1; background:rgba(127,127,127,.18); }
.sheetedit-slicer-items { flex:1; overflow:auto; padding:4px; display:grid; gap:3px; }
.sheetedit-slicer-item { border:1px solid var(--sheetedit-border,#c8ccd2); border-radius:4px;
background:var(--se-slicer-off-bg,transparent); color:var(--se-slicer-off-fg,inherit);
font:inherit; padding:3px 6px; cursor:pointer;
text-align:left; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; opacity:.45; }
.sheetedit-slicer-item:hover { border-color:var(--sheetedit-accent,#4c8bf5); }
.sheetedit-slicer-item.on { opacity:1; background:var(--se-slicer-accent,var(--sheetedit-accent,#4c8bf5));
color:var(--se-slicer-on-fg,#fff);
border-color:var(--se-slicer-accent,var(--sheetedit-accent,#4c8bf5)); }
.sheetedit-slicerbox.styled .sheetedit-slicer-item { opacity:1; }
.sheetedit-slicerbox.readonly .sheetedit-slicer-item { cursor:default; }
.sheetedit-slicerbox.readonly .sheetedit-slicer-clear { display:none; }
.sheetedit-imagelayer { position:absolute; overflow:hidden; pointer-events:none; z-index:5; }
.sheetedit-imagelayer-inner { position:absolute; inset:0; }
.sheetedit-imagebox { position:absolute; }
.sheetedit-imagebox img { width:100%; height:100%; object-fit:contain; display:block; pointer-events:none; }
.sheetedit-imagebox.editable { pointer-events:auto; cursor:move; }
.sheetedit-imagebox.selected { outline:1.5px solid var(--sheetedit-accent,#4c8bf5); outline-offset:1px; }
.sheetedit-image-resize { position:absolute; right:-5px; bottom:-5px; width:12px; height:12px; border-radius:3px;
background:var(--sheetedit-accent,#4c8bf5); border:1.5px solid var(--sheetedit-handle-border); cursor:nwse-resize; pointer-events:auto; display:none; }
.sheetedit-imagebox.selected .sheetedit-image-resize { display:block; }
.sheetedit-pivotlayer { position:absolute; overflow:hidden; pointer-events:none; z-index:4; }
.sheetedit-pivotlayer-inner { position:absolute; inset:0; }
.sheetedit-pivotbox { position:absolute; box-sizing:border-box; border:1.5px dashed var(--sheetedit-accent,#3b82f6); border-radius:3px; background:color-mix(in srgb, var(--sheetedit-accent,#3b82f6) 6%, transparent); }
.sheetedit-pivottag { position:absolute; top:0; left:0; transform:translateY(-100%); pointer-events:auto; font:600 10px/1.4 system-ui,sans-serif; color:var(--sheetedit-accent-fg); background:var(--sheetedit-accent,#3b82f6); padding:1px 6px; border-radius:3px 3px 0 0; white-space:nowrap; cursor:default; }
.sheetedit-shapelayer { position:absolute; overflow:hidden; pointer-events:none; z-index:5; }
.sheetedit-shapelayer-inner { position:absolute; inset:0; }
.sheetedit-shapebox { position:absolute; }
.sheetedit-shapebox svg { width:100%; height:100%; display:block; overflow:visible; pointer-events:none; }
.sheetedit-shapebox.editable { pointer-events:auto; cursor:move; }
.sheetedit-shapebox.selected { outline:1.5px dashed var(--sheetedit-accent,#4c8bf5); outline-offset:2px; }
.sheetedit-shape-resize { position:absolute; right:-5px; bottom:-5px; width:12px; height:12px; border-radius:3px;
background:var(--sheetedit-accent,#4c8bf5); border:1.5px solid var(--sheetedit-handle-border); cursor:nwse-resize; pointer-events:auto; display:none; }
.sheetedit-shapebox.selected .sheetedit-shape-resize { display:block; }
.sheetedit-shape-del { position:absolute; right:-9px; top:-9px; width:16px; height:16px; border-radius:50%; padding:0;
background:var(--sheetedit-danger-strong); color:var(--sheetedit-accent-fg); border:1.5px solid var(--sheetedit-handle-border); cursor:pointer; pointer-events:auto; display:none;
font:700 11px/13px sans-serif; text-align:center; }
.sheetedit-shapebox.selected .sheetedit-shape-del { display:block; }
.sheetedit-outline { position:absolute; overflow:hidden; z-index:12;
background:var(--sheetedit-chrome,#2b2f36); border-right:1px solid var(--sheetedit-border,#1c1f24); }
.sheetedit-outline-inner { position:absolute; left:0; right:0; top:0; }
.sheetedit-outline-head { position:absolute; left:0; top:0; display:flex; align-items:center;
justify-content:flex-start; gap:1px; padding:0 1px;
background:var(--sheetedit-chrome,#2b2f36); border-bottom:1px solid var(--sheetedit-border,#1c1f24); z-index:2; }
.sheetedit-outline-lvl { width:12px; height:12px; padding:0; line-height:1; font:9px system-ui,sans-serif;
border:1px solid var(--sheetedit-border,#4a4f57); border-radius:2px; cursor:pointer;
background:var(--sheetedit-btn,#3a3f47); color:var(--sheetedit-text,#e6e6e6); }
.sheetedit-outline-lvl:hover { border-color:var(--sheetedit-accent,#6e7bff); }
.sheetedit-outline-bar { position:absolute; width:1px; background:var(--sheetedit-muted,#8b93a1); }
.sheetedit-outline-foot { position:absolute; height:1px; background:var(--sheetedit-muted,#8b93a1); }
.sheetedit-outline-btn { position:absolute; width:11px; height:11px; padding:0; line-height:9px;
font:9px/9px system-ui,sans-serif; text-align:center; cursor:pointer;
border:1px solid var(--sheetedit-muted,#8b93a1); border-radius:2px;
background:var(--sheetedit-chrome,#2b2f36); color:var(--sheetedit-text,#e6e6e6); }
.sheetedit-outline-btn:hover { border-color:var(--sheetedit-accent,#6e7bff); }
.sheetedit-timelinelayer { position:absolute; overflow:hidden; pointer-events:none; z-index:6; }
.sheetedit-timelinelayer-inner { position:absolute; inset:0; }
.sheetedit-timelinebox { position:absolute; pointer-events:auto; display:flex; flex-direction:column;
background:var(--sheetedit-chrome,#fff); color:var(--sheetedit-text,#1c1f24);
border:1px solid var(--sheetedit-border,#c8ccd2); border-radius:6px;
box-shadow:0 2px 10px rgba(0,0,0,.18); font:12px system-ui,sans-serif; overflow:hidden; }
.sheetedit-timeline-head { display:flex; align-items:baseline; gap:6px; padding:5px 7px;
border-bottom:1px solid var(--sheetedit-border,#c8ccd2); background:rgba(127,127,127,.08); }
.sheetedit-timeline-title { font-weight:600; flex:0 0 auto; }
.sheetedit-timeline-range { flex:1; color:var(--sheetedit-muted,#6b7280); overflow:hidden;
text-overflow:ellipsis; white-space:nowrap; font-size:11px; }
.sheetedit-timeline-clear { border:0; background:none; cursor:pointer; color:inherit; opacity:.65;
font:inherit; padding:1px 4px; border-radius:4px; }
.sheetedit-timeline-clear:hover { opacity:1; background:rgba(127,127,127,.18); }
.sheetedit-timeline-periods { flex:1; display:flex; align-items:stretch; gap:2px; padding:5px; overflow-x:auto; }
.sheetedit-timeline-period { flex:1 0 34px; border:1px solid var(--sheetedit-border,#c8ccd2);
border-radius:3px; background:transparent; color:inherit; font:inherit; font-size:11px;
cursor:pointer; padding:2px; opacity:.45; white-space:nowrap; }
.sheetedit-timeline-period.on { opacity:1; background:var(--sheetedit-accent,#4c8bf5); color:var(--sheetedit-accent-fg);
border-color:var(--sheetedit-accent,#4c8bf5); }
.sheetedit-chart-resize { position:absolute; right:-3px; bottom:-3px; width:12px; height:12px; cursor:nwse-resize; z-index:2;
background:var(--sheetedit-accent, #6e7bff); border:2px solid var(--sheetedit-handle-border); border-radius:3px; opacity:0; }
.sheetedit-chartbox.sel .sheetedit-chart-resize { opacity:1; }
.sheetedit-chartbox { cursor:move; }
.sheetedit-chart-modal { position:fixed; inset:0; z-index:60; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,.45); padding:16px; }
.sheetedit-chart-card { display:flex; flex-direction:column; width:min(860px,100%); max-height:90vh; background:var(--sheetedit-chrome, #2b2f36); color:var(--sheetedit-text, #e6e6e6);
border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 14px 44px rgba(0,0,0,.5); font:13px system-ui,sans-serif; overflow:hidden; }
.sheetedit-chart-head { display:flex; align-items:center; justify-content:space-between; padding:14px 16px; border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.sheetedit-chart-head h3 { margin:0; font-size:15px; }
.sheetedit-chart-x { font:inherit; font-size:18px; line-height:1; width:28px; height:28px; border:none; border-radius:6px; background:transparent; color:var(--sheetedit-muted, #aab2bf); cursor:pointer; }
.sheetedit-chart-x:hover { background:var(--sheetedit-btn, #3a3f47); color:inherit; }
.sheetedit-chart-body { display:flex; gap:18px; padding:16px; overflow:hidden; min-height:0; }
.sheetedit-chart-opts { flex:1 1 auto; min-width:0; overflow-y:auto; padding-right:4px; }
.sheetedit-chart-side { flex:0 0 340px; display:flex; flex-direction:column; }
.sheetedit-chart-sec { margin-bottom:14px; }
.sheetedit-chart-sec[hidden] { display:none; }
.sheetedit-chart-card h4 { margin:0 0 8px; font-size:11px; text-transform:uppercase; letter-spacing:.05em; color:var(--sheetedit-muted, #aab2bf); border-bottom:1px solid var(--sheetedit-border, #1c1f24); padding-bottom:4px; }
.sheetedit-chart-types { display:flex; flex-wrap:wrap; gap:6px; }
.sheetedit-chart-type { font:inherit; font-size:12px; padding:6px 11px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
.sheetedit-chart-type.sel { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); }
.sheetedit-chart-field { display:flex; flex-direction:column; gap:4px; margin-bottom:10px; }
.sheetedit-chart-field[hidden] { display:none; }
.sheetedit-chart-field > span { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
.sheetedit-chart-field input[type=text], .sheetedit-chart-field select, .sheetedit-chart-row select, .sheetedit-chart-row input[type=number] {
font:inherit; background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); padding:6px 8px; }
.sheetedit-chart-checks { display:flex; flex-wrap:wrap; gap:12px 16px; font-size:13px; }
.sheetedit-chart-checks label { display:flex; align-items:center; gap:6px; }
.sheetedit-chart-row { display:flex; flex-wrap:wrap; align-items:center; gap:12px 16px; margin-top:8px; }
.sheetedit-chart-row label { display:flex; align-items:center; gap:6px; font-size:12px; color:var(--sheetedit-muted, #aab2bf); }
.sheetedit-chart-row input[type=number] { width:64px; }
.sheetedit-chart-two { display:flex; gap:12px; }
.sheetedit-chart-two > * { flex:1; }
.sheetedit-chart-swatches { display:flex; flex-wrap:wrap; gap:10px 14px; }
.sheetedit-chart-swatch { display:flex; align-items:center; gap:6px; font-size:12px; }
.sheetedit-chart-swatch input[type=color] { width:26px; height:22px; padding:0; border:1px solid var(--sheetedit-btn, #3a4047); border-radius:4px; background:none; cursor:pointer; }
.sheetedit-chart-fill { display:flex; align-items:center; gap:6px; }
.sheetedit-chart-preview { flex:1 1 auto; min-height:240px; background:var(--sheetedit-cell-bg); border-radius:6px; padding:8px; }
.sheetedit-chart-foot { display:flex; justify-content:flex-end; gap:8px; padding:12px 16px; border-top:1px solid var(--sheetedit-border, #1c1f24); }
.sheetedit-chart-btn { font:inherit; font-size:13px; padding:7px 16px; border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:6px;
background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6); cursor:pointer; }
.sheetedit-chart-btn.primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); }
@media (max-width:720px) {
.sheetedit-chart-body { flex-direction:column; overflow-y:auto; }
.sheetedit-chart-side { flex-basis:auto; order:-1; }
.sheetedit-chart-preview { min-height:170px; height:170px; }
.sheetedit-chart-opts { overflow:visible; }
}
.sheetedit-chart-editbar { position:fixed; z-index:30; display:flex; align-items:center; gap:6px; padding:5px 7px;
background:var(--sheetedit-chrome, #2b2f36); border:1px solid var(--sheetedit-border, #1c1f24); border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,.4); }
.sheetedit-chart-editbar[hidden] { display:none; }
.sheetedit-chart-editbar button { font:inherit; font-size:12px; background:var(--sheetedit-btn, #3a3f47);
color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; padding:4px 10px; cursor:pointer; }
.sheetedit-panediv-layer { position:absolute; overflow:hidden; pointer-events:none; z-index:11; }
.sheetedit-panediv { position:absolute; pointer-events:auto; background:var(--sheetedit-accent,#6e7bff); opacity:.55; }
.sheetedit-panediv:hover, .sheetedit-panediv.dragging { opacity:1; }
.sheetedit-panediv-h { left:0; right:0; height:3px; cursor:row-resize; }
.sheetedit-panediv-v { top:0; bottom:0; width:3px; cursor:col-resize; }
.sheetedit-panediv::after { content:""; position:absolute; inset:-4px; }
.se-pqe { position:fixed; inset:0; z-index:60; display:flex; flex-direction:column;
background:var(--sheetedit-bg, #1f2227); color:var(--sheetedit-text, #e6e6e6);
font:13px system-ui, sans-serif; }
.se-pqe[hidden] { display:none; }
.se-pqe-bar { display:flex; align-items:center; gap:10px; padding:8px 12px;
background:var(--sheetedit-chrome, #2b2f36); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-title { font-weight:600; }
.se-pqe-spacer { flex:1; }
.se-pqe-btn { font:inherit; font-size:13px; display:inline-flex; align-items:center; gap:6px; background:var(--sheetedit-btn, #3a3f47);
color:var(--sheetedit-text, #e6e6e6); border:1px solid var(--sheetedit-btn-border, #4a4f57);
border-radius:6px; padding:5px 12px; cursor:pointer; }
.se-pqe-btn svg { display:block; }
.se-pqe-btn:hover:not(:disabled) { background:var(--sheetedit-btn-hover, #454b54); }
.se-pqe-btn:disabled { opacity:.5; cursor:default; }
.se-pqe-btn.primary { background:var(--sheetedit-accent, #6e7bff); border-color:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); }
.se-pqe-ribbon { display:flex; align-items:stretch; padding:6px 2px 3px; overflow-x:auto;
background:var(--sheetedit-chrome, #2b2f36); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-grp { display:flex; flex-direction:column; padding:0 9px; border-right:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-grp:last-child { border-right:0; }
.se-pqe-grp-body { display:grid; grid-template-rows:repeat(3, 23px); grid-auto-flow:column; grid-auto-columns:max-content; gap:2px 6px; }
.se-pqe-grp-label { text-align:center; font-size:10.5px; color:var(--sheetedit-muted, #aab2bf); padding-top:4px; margin-top:auto; }
.se-pqe-rbtn { font:inherit; font-size:12px; display:inline-flex; align-items:center; gap:8px; width:100%; justify-content:flex-start;
background:transparent; color:var(--sheetedit-text, #e6e6e6); border:1px solid transparent; border-radius:5px; padding:0 9px 0 7px; cursor:pointer; white-space:nowrap; }
.se-pqe-rbtn svg { display:block; flex:none; width:16px; height:16px; color:var(--sheetedit-accent, #6e7bff); }
.se-pqe-rbtn:hover:not(:disabled) { background:var(--sheetedit-btn, #3a3f47); }
.se-pqe-rbtn:disabled { opacity:.4; cursor:default; }
.se-pqe-rbtn:disabled svg { color:var(--sheetedit-muted, #aab2bf); }
.se-pqe-modal { position:absolute; inset:0; z-index:5; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,.45); }
.se-pqe-modal[hidden] { display:none; }
.se-pqe-card { width:min(420px,92%); max-height:80%; overflow:auto; background:var(--sheetedit-chrome, #2b2f36);
border:1px solid var(--sheetedit-border, #1c1f24); border-radius:10px; box-shadow:0 12px 40px rgba(0,0,0,.5); padding:16px; }
.se-pqe-card h3 { margin:0 0 12px; font-size:15px; }
.se-pqe-field { display:flex; flex-direction:column; gap:4px; margin-bottom:11px; }
.se-pqe-field > span { color:var(--sheetedit-muted, #aab2bf); font-size:12px; }
.se-pqe-field input, .se-pqe-field select { font:inherit; background:var(--sheetedit-border, #1c1f24);
border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; color:var(--sheetedit-text, #e7eaf0); padding:5px 8px; }
.se-pqe-checks { display:flex; flex-direction:column; gap:3px; max-height:180px; overflow:auto;
border:1px solid var(--sheetedit-btn, #3a4047); border-radius:5px; padding:6px 8px; background:var(--sheetedit-border, #1c1f24); }
.se-pqe-checks label { display:flex; align-items:center; gap:7px; font-size:13px; }
.se-pqe-card-actions { display:flex; justify-content:flex-end; gap:8px; margin-top:6px; }
.se-pqe-fx { display:flex; align-items:stretch; gap:6px; padding:6px 12px;
background:var(--sheetedit-chrome2, #23262c); border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-fx-lbl { align-self:center; color:var(--sheetedit-muted, #aab2bf); font:12px ui-monospace,monospace; }
.se-pqe-fx textarea { flex:1; min-height:26px; max-height:120px; resize:vertical;
background:var(--sheetedit-border, #1c1f24); border:1px solid var(--sheetedit-btn, #3a4047);
border-radius:5px; color:var(--sheetedit-text, #e7eaf0); font:13px ui-monospace,monospace; padding:5px 8px; }
.se-pqe-main { flex:1; min-height:0; display:flex; position:relative; }
.se-pqe-queries { width:200px; flex:none; overflow:auto; border-right:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-settings { width:240px; flex:none; overflow:auto; border-left:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-panetoggle { display:none; }
.se-pqe-center { flex:1; min-width:0; display:flex; flex-direction:column; }
.se-pqe-pane-h { padding:7px 12px; font-weight:600; color:var(--sheetedit-muted, #aab2bf);
font-size:12px; text-transform:uppercase; letter-spacing:.04em; border-bottom:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-pane-h-row { display:flex; align-items:center; justify-content:space-between; }
.se-pqe-newq { display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; padding:0; cursor:pointer;
background:var(--sheetedit-btn, #3a3f47); color:var(--sheetedit-text, #e6e6e6);
border:1px solid var(--sheetedit-btn-border, #4a4f57); border-radius:5px; }
.se-pqe-newq svg { display:block; }
.se-pqe-newq:hover { background:var(--sheetedit-btn-hover, #454b54); }
.se-pqe-item { display:flex; align-items:center; gap:6px; padding:7px 12px; cursor:pointer; border-bottom:1px solid rgba(0,0,0,.12); }
.se-pqe-item:hover { background:var(--sheetedit-btn, #3a3f47); }
.se-pqe-item.sel { background:var(--sheetedit-accent, #6e7bff); color:var(--sheetedit-accent-fg); }
.se-pqe-item-name { flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.se-pqe-item-x { opacity:0; border:0; background:none; color:inherit; cursor:pointer; font-size:15px; line-height:1; padding:0 2px; border-radius:4px; }
.se-pqe-item:hover .se-pqe-item-x { opacity:.7; }
.se-pqe-item-x:hover { opacity:1 !important; background:rgba(255,255,255,.15); }
.se-pqe-name-in { width:100%; box-sizing:border-box; font:inherit; background:var(--sheetedit-border, #1c1f24);
border:1px solid var(--sheetedit-accent, #6e7bff); border-radius:4px; color:inherit; padding:3px 6px; }
.se-pqe-preview { flex:1; min-height:0; overflow:auto; background:var(--sheetedit-grid-bg); }
.se-pqe-ptable { border-collapse:collapse; font:13px/1.3 ui-sans-serif, system-ui, sans-serif; color:var(--sheetedit-cell-fg); }
.se-pqe-ptable th, .se-pqe-ptable td { border:1px solid var(--sheetedit-head-border); padding:3px 9px; text-align:left; white-space:nowrap; }
.se-pqe-ptable th { position:sticky; top:0; background:var(--sheetedit-head-bg); color:var(--sheetedit-head-fg); font-weight:600; z-index:1; }
.se-pqe-ptable th .nm { display:block; }
.se-pqe-ptable th .ty { display:block; font-weight:400; font-size:10px; color:var(--sheetedit-faint); }
.se-pqe-ptable th .qbar { display:flex; height:3px; margin-top:3px; border-radius:2px; overflow:hidden; background:var(--sheetedit-head-border); }
.se-pqe-ptable th .qv { background:var(--sheetedit-ok); }
.se-pqe-ptable th .qe { background:var(--sheetedit-neutral); }
.se-pqe-ptable th .qx { background:var(--sheetedit-danger); }
.se-pqe-ptable td.num { text-align:right; font-variant-numeric:tabular-nums; }
.se-pqe-ptable td.null, .se-pqe-ptable td.obj { color:var(--sheetedit-faint); }
.se-pqe-ptable td.err { color:var(--sheetedit-danger); }
.se-pqe-ptable tr:nth-child(even) td { background:var(--sheetedit-zebra); }
.se-pqe-scalar { padding:16px; font:14px ui-monospace,monospace; color:var(--sheetedit-text, #e6e6e6); }
.se-pqe-foot { display:flex; align-items:center; gap:14px; padding:5px 12px; color:var(--sheetedit-muted, #aab2bf);
font-size:12px; background:var(--sheetedit-chrome2, #23262c); border-top:1px solid var(--sheetedit-border, #1c1f24); }
.se-pqe-foot .err { color:var(--sheetedit-error-text); }
.se-pqe-empty { padding:24px; color:var(--sheetedit-muted, #aab2bf); }
@media (max-width: 760px) {
.se-pqe-title { display:none; }
.se-pqe-bar { gap:6px; padding:7px 8px; }
.se-pqe-bar .se-pqe-btn span { display:none; }
.se-pqe-btn { padding:6px 8px; }
.se-pqe-panetoggle { display:inline-flex; }
.se-pqe-rbtn span { display:none; }
.se-pqe-rbtn { gap:0; padding:0 7px; justify-content:center; width:auto; }
.se-pqe-grp { padding:0 6px; }
.se-pqe-queries, .se-pqe-settings {
position:absolute; top:0; bottom:0; z-index:6; width:min(280px, 82%);
background:var(--sheetedit-chrome, #2b2f36); box-shadow:0 0 30px rgba(0,0,0,.45);
transition:transform .18s ease; will-change:transform;
}
.se-pqe-queries { left:0; border-right:1px solid var(--sheetedit-border, #1c1f24); transform:translateX(-102%); }
.se-pqe-settings { right:0; border-left:1px solid var(--sheetedit-border, #1c1f24); transform:translateX(102%); }
.se-pqe.show-queries .se-pqe-queries { transform:none; }
.se-pqe.show-steps .se-pqe-settings { transform:none; }
.se-pqe.show-queries .se-pqe-center::after, .se-pqe.show-steps .se-pqe-center::after {
content:""; position:absolute; inset:0; z-index:5; background:rgba(0,0,0,.35);
}
.se-pqe-center { position:relative; }
}
.sheetedit-measure {
position:absolute; visibility:hidden; left:-9999px; top:0;
white-space:pre-wrap; word-break:break-word; padding:3px 8px; box-sizing:border-box; line-height:1.3; font:inherit;
}
.sheetedit-theme-list { display:flex; flex-direction:column; gap:6px; margin-bottom:12px; }
.sheetedit-theme-opt {
display:flex; align-items:center; gap:10px; padding:7px 9px; border-radius:7px; cursor:pointer;
background:none; color:inherit; font:inherit; text-align:left;
border:1px solid var(--sheetedit-btn-border);
}
.sheetedit-theme-opt:hover { background:var(--sheetedit-btn); }
.sheetedit-theme-opt.is-current { border-color:var(--sheetedit-accent); }
.sheetedit-theme-name { flex:1 1 auto; }
.sheetedit-theme-tag { font-size:11px; color:var(--sheetedit-muted); }
.sheetedit-theme-swatches { display:flex; gap:2px; flex:none; }
.sheetedit-theme-chip { width:13px; height:16px; border-radius:2px; border:1px solid var(--sheetedit-border); }
.sheetedit-print { position:absolute; left:-100000px; top:0; }
.sheetedit-print-page { position:relative; overflow:hidden; background:var(--sheetedit-paper); color:var(--sheetedit-ink); box-sizing:border-box; break-after:page; }
.sheetedit-print-page:last-child { break-after:auto; }
.sheetedit-print-body { position:absolute; overflow:hidden; display:flex; }
.sheetedit-print-band {
position:absolute; display:flex; align-items:center; font:11px system-ui, sans-serif; color:var(--sheetedit-ink);
}
.sheetedit-print-band > span { flex:1 1 0; min-width:0; }
.sheetedit-print-left { text-align:left; }
.sheetedit-print-center { text-align:center; }
.sheetedit-print-right { text-align:right; }
.sheetedit-print-table { border-collapse:collapse; table-layout:fixed; font:13px system-ui, sans-serif; color:var(--sheetedit-ink); }
.sheetedit-print-table td { padding:1px 4px; overflow:hidden; vertical-align:bottom; word-break:break-word; }
.sheetedit-print-table td.num { text-align:right; }
.sheetedit-print-table.has-grid td, .sheetedit-print-table.has-grid th { border:1px solid var(--sheetedit-print-rule); }
.sheetedit-print-head {
border:1px solid var(--sheetedit-print-rule); background:var(--sheetedit-print-head-bg); font-weight:600; font-size:11px; text-align:center; padding:1px 4px;
}
@media print {
body > *:not(.sheetedit-print) { display:none !important; }
.sheetedit-print { position:static !important; left:auto !important; }
.sheetedit-print-page { break-after:page; }
}
.sheetedit-ctrllayer { position:absolute; overflow:hidden; pointer-events:none; z-index:14; }
.sheetedit-ctrllayer-inner { position:absolute; left:0; top:0; }
.sheetedit-ctrlbox {
position:absolute; pointer-events:auto; display:flex; align-items:center; overflow:hidden;
font:12px system-ui, sans-serif; color:var(--sheetedit-cell-fg);
}
.sheetedit-ctrlbox > * { max-width:100%; }
.sheetedit-ctrl-check { display:flex; align-items:center; gap:5px; cursor:pointer; white-space:nowrap; }
.sheetedit-ctrl-check input { margin:0; cursor:pointer; }
.sheetedit-ctrl-select, .sheetedit-ctrl-num {
font:inherit; width:100%; box-sizing:border-box; padding:1px 4px; border-radius:3px;
background:var(--sheetedit-cell-bg); color:var(--sheetedit-cell-fg); border:1px solid var(--sheetedit-head-border);
}
.sheetedit-ctrl-button {
font:inherit; width:100%; height:100%; border-radius:4px; cursor:not-allowed;
background:var(--sheetedit-flat-bg); color:var(--sheetedit-flat-fg); border:1px solid var(--sheetedit-head-border);
}
.sheetedit-ctrl-label { white-space:nowrap; }
.sheetedit-ctrl-group {
width:100%; height:100%; box-sizing:border-box; padding:1px 4px;
border:1px solid var(--sheetedit-head-border); border-radius:3px; color:var(--sheetedit-head-fg);
}`;
