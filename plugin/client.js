window.__ModuleLoader__.load({
  id: '@local/dsh-word-pipeline',
  factory(require) {
    const React = require('react');
    const h = React.createElement;
    const API = '/word-pipeline';
    const token = () => (window.__DSH_WORD_PIPELINE__ || {}).token || '';
    let runtime = null;
    let uiWorkspace = null;
    let clientSessions = null;

    const CSS = `
.wp-root{display:flex;flex-direction:column;gap:12px;padding:12px;height:100%;box-sizing:border-box;
  color:var(--dsw-alias-label-primary);background:var(--dsw-alias-bg-base);font-size:13px;overflow:auto}
.wp-head{display:flex;align-items:center;justify-content:space-between;gap:8px}
.wp-title{font-size:14px;font-weight:600}
.wp-badge{font-size:11px;padding:1px 8px;border-radius:999px;color:var(--dsw-alias-label-secondary);
  border:1px solid var(--dsw-alias-border-l1)}
.wp-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.wp-drop{display:flex;flex-direction:column;gap:4px;padding:10px;border-radius:10px;cursor:pointer;
  border:1px dashed var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-1);min-height:64px}
.wp-drop[data-over="1"]{border-color:var(--dsw-alias-brand-primary)}
.wp-drop-label{font-size:12px;font-weight:600}
.wp-drop-hint{font-size:11px;color:var(--dsw-alias-label-secondary);word-break:break-all;line-height:1.4}
.wp-actions{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.wp-btn{appearance:none;border:1px solid var(--dsw-alias-border-l2);background:var(--dsw-alias-bg-layer-1);
  color:var(--dsw-alias-label-primary);border-radius:9px;padding:8px 10px;font-size:13px;cursor:pointer}
.wp-btn:hover{border-color:var(--dsw-alias-brand-primary)}
.wp-btn[disabled]{opacity:.5;cursor:default}
.wp-btn[data-busy="1"]{border-color:var(--dsw-alias-brand-primary)}
.wp-status{font-size:12px;color:var(--dsw-alias-label-secondary);display:flex;gap:8px;align-items:center}
.wp-dot{width:8px;height:8px;border-radius:50%;background:var(--dsw-alias-state-idle-primary)}
.wp-dot[data-state="ok"]{background:var(--dsw-alias-state-success-primary)}
.wp-dot[data-state="bad"]{background:var(--dsw-alias-state-error-primary)}
.wp-dot[data-state="run"]{background:var(--dsw-alias-state-warn-primary)}
.wp-log{margin:0;padding:10px;border-radius:10px;background:var(--dsw-alias-bg-layer-2);font-size:11.5px;
  line-height:1.5;white-space:pre-wrap;word-break:break-word;max-height:320px;overflow:auto;
  font-family:ui-monospace,SFMono-Regular,Consolas,monospace}
.wp-error{color:var(--dsw-alias-state-error-primary)}
.wp-files{display:flex;flex-direction:column;gap:4px;font-size:12px;color:var(--dsw-alias-label-secondary)}
.wp-path{font-family:ui-monospace,Consolas,monospace;font-size:11px;word-break:break-all}
.wp-cmd{font-family:ui-monospace,Consolas,monospace;font-size:11px;color:var(--dsw-alias-label-secondary);
  padding:6px 8px;border-radius:8px;background:var(--dsw-alias-bg-layer-1);word-break:break-all}
.wp-launch{appearance:none;border:1px solid var(--dsw-alias-border-l1);background:transparent;
  color:var(--dsw-alias-label-secondary);border-radius:999px;padding:4px 10px;font-size:12px;cursor:pointer}
.wp-launch:hover{color:var(--dsw-alias-label-primary);border-color:var(--dsw-alias-border-l2)}
.wp-section{display:flex;flex-direction:column;gap:6px}
.wp-section-title{font-size:12px;font-weight:600;display:flex;justify-content:space-between}
.wp-steps{display:flex;flex-direction:column;gap:4px;max-height:132px;overflow:auto}
.wp-step{appearance:none;text-align:left;border:1px solid var(--dsw-alias-border-l1);background:var(--dsw-alias-bg-layer-1);
  color:var(--dsw-alias-label-secondary);border-radius:8px;padding:4px 8px;font-size:11.5px;cursor:pointer;
  font-family:ui-monospace,Consolas,monospace;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.wp-step[data-active="1"]{border-color:var(--dsw-alias-brand-primary);color:var(--dsw-alias-label-primary)}
.wp-stage{border:1px solid var(--dsw-alias-border-l1);border-radius:10px;background:var(--dsw-alias-bg-layer-2);
  min-height:180px;display:flex;align-items:center;justify-content:center;overflow:hidden}
.wp-stage img{width:100%;height:auto;max-height:none;object-fit:contain;display:block;margin:0 auto}
.wp-stage-note{font-size:11px;color:var(--dsw-alias-label-secondary);padding:16px;text-align:center}
.wp-page{--wpa-accent:var(--dsw-alias-brand-primary);--wpa-line:var(--dsw-alias-border-l1);
  --wpa-strong:var(--dsw-alias-border-l2);--wpa-muted:var(--dsw-alias-label-secondary);
  --wpa-surface:var(--dsw-alias-bg-layer-1);--wpa-surface2:var(--dsw-alias-bg-layer-2);
  height:100%;overflow:auto;color:var(--dsw-alias-label-primary);
  background:radial-gradient(900px 320px at 12% -8%,
    color-mix(in oklab, var(--wpa-accent) 12%, transparent), transparent 62%),var(--dsw-alias-bg-base)}
.wp-page .wp-page-inner{max-width:1180px;margin:0 auto;padding:30px 26px 44px;display:flex;flex-direction:column;
  gap:20px;min-height:100%;box-sizing:border-box}
.wp-page .wp-hero{display:flex;align-items:center;justify-content:space-between;gap:16px}
.wp-page .wp-hero-left{display:flex;align-items:center;gap:14px;min-width:0}
.wp-page .wp-logo{width:48px;height:48px;border-radius:14px;display:grid;place-items:center;flex:none;
  color:var(--wpa-accent);background:color-mix(in oklab, var(--wpa-accent) 12%, var(--wpa-surface));
  border:1px solid color-mix(in oklab, var(--wpa-accent) 26%, transparent);
  box-shadow:0 8px 22px color-mix(in oklab, var(--wpa-accent) 16%, transparent)}
.wp-page .wp-hero-title{font-size:20px;font-weight:600;letter-spacing:0;line-height:1.25}
.wp-page .wp-hero-sub{font-size:12.5px;color:var(--wpa-muted);margin-top:3px;line-height:1.5}
.wp-page .wp-pill{display:inline-flex;align-items:center;gap:6px;font-size:11.5px;color:var(--wpa-muted);
  border:1px solid var(--wpa-line);background:var(--wpa-surface);padding:5px 11px;border-radius:999px}
.wp-page .wp-pill-dot{width:7px;height:7px;border-radius:50%;background:var(--dsw-alias-state-success-primary)}
.wp-page .wp-list{display:flex;flex-direction:column;gap:10px}
.wp-page .wp-row{display:flex;align-items:center;gap:16px;padding:18px 20px;border:1px solid var(--wpa-line);
  border-radius:16px;background:var(--wpa-surface);cursor:pointer;text-align:left;color:inherit;font:inherit;
  transition:border-color .16s ease, box-shadow .16s ease, transform .16s ease}
.wp-page .wp-row:hover{border-color:color-mix(in oklab, var(--wpa-accent) 45%, var(--wpa-line));
  box-shadow:0 12px 28px rgba(0,0,0,.10);transform:translateY(-1px)}
.wp-page .wp-row-icon{width:44px;height:44px;border-radius:12px;display:grid;place-items:center;flex:none;
  color:var(--wpa-accent);background:color-mix(in oklab, var(--wpa-accent) 10%, var(--wpa-surface));
  border:1px solid color-mix(in oklab, var(--wpa-accent) 22%, transparent)}
.wp-page .wp-row-text{flex:1;min-width:0;display:flex;flex-direction:column;gap:4px}
.wp-page .wp-row-title{font-weight:600;font-size:14px}
.wp-page .wp-row-desc{font-size:12.5px;color:var(--wpa-muted);line-height:1.5}
.wp-page .wp-row-meta{font-size:11px;color:var(--wpa-muted);border:1px solid var(--wpa-line);
  border-radius:999px;padding:3px 9px;white-space:nowrap}
.wp-page .wp-row-arrow{font-size:16px;color:var(--wpa-muted)}
.wp-page .wp-row-note{font-size:11.5px;color:var(--wpa-muted);padding:2px 4px;line-height:1.6}
.wp-page .wp-workspace{display:grid;grid-template-columns:minmax(0,.94fr) minmax(0,1.06fr);gap:16px;
  flex:1;min-height:480px;width:100%;align-items:stretch}
.wp-page .wp-panel{display:flex;flex-direction:column;min-width:0;min-height:0;border:1px solid var(--wpa-line);
  border-radius:18px;background:var(--wpa-surface);overflow:hidden}
.wp-page .wp-panel-head{display:flex;align-items:center;justify-content:space-between;gap:10px;
  padding:12px 16px;border-bottom:1px solid var(--wpa-line)}
.wp-page .wp-panel-title{font-size:13px;font-weight:600;display:flex;align-items:center;gap:8px}
.wp-page .wp-panel-meta{font-size:11.5px;color:var(--wpa-muted)}
.wp-page .wp-panel-body{flex:1;min-height:0;display:flex;flex-direction:column;gap:12px;padding:14px 16px 16px;overflow:auto}
.wp-page .wp-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.wp-page .wp-drop{display:flex;align-items:center;gap:11px;padding:13px;border:1px dashed var(--wpa-strong);
  border-radius:14px;background:var(--wpa-surface2);cursor:pointer;min-height:78px;
  transition:border-color .16s ease, background .16s ease}
.wp-page .wp-drop:hover{border-color:var(--wpa-accent);
  background:color-mix(in oklab, var(--wpa-accent) 6%, var(--wpa-surface2))}
.wp-page .wp-drop[data-over="1"]{border-style:solid;border-color:var(--wpa-accent)}
.wp-page .wp-drop-icon{width:34px;height:34px;border-radius:10px;display:grid;place-items:center;flex:none;
  font-size:16px;color:var(--wpa-accent);background:color-mix(in oklab, var(--wpa-accent) 10%, var(--wpa-surface));
  border:1px solid color-mix(in oklab, var(--wpa-accent) 20%, transparent)}
.wp-page .wp-drop-text{display:flex;flex-direction:column;gap:3px;min-width:0}
.wp-page .wp-drop-label{font-size:12.5px;font-weight:600}
.wp-page .wp-drop-hint{font-size:11px;color:var(--wpa-muted);line-height:1.4;word-break:break-all}
.wp-page .wp-chat{flex:1;min-height:150px;overflow:auto;display:flex;flex-direction:column;gap:10px;padding:2px}
.wp-page .wp-msg{max-width:88%;padding:9px 12px;border-radius:14px;font-size:12.5px;line-height:1.55;
  white-space:pre-wrap;word-break:break-word}
.wp-page .wp-msg[data-role="user"]{align-self:flex-end;background:var(--wpa-accent);color:#fff;
  border-bottom-right-radius:5px;box-shadow:0 6px 16px color-mix(in oklab, var(--wpa-accent) 28%, transparent)}
.wp-page .wp-msg[data-role="agent"]{align-self:flex-start;background:var(--wpa-surface2);
  border:1px solid var(--wpa-line);border-bottom-left-radius:5px}
.wp-page .wp-chip{border:1px solid var(--wpa-line);background:var(--wpa-surface);color:var(--wpa-muted);
  border-radius:999px;padding:5px 11px;font-size:11.5px;cursor:pointer;transition:color .15s, border-color .15s}
.wp-page .wp-chip:hover{color:var(--wpa-accent);border-color:var(--wpa-accent)}
.wp-page .wp-chip[disabled]{opacity:.5;cursor:default}
.wp-page .wp-input-row{display:flex;gap:8px}
.wp-page .wp-input{flex:1;min-width:0;border:1px solid var(--wpa-strong);background:var(--dsw-alias-bg-base);
  color:var(--dsw-alias-label-primary);border-radius:12px;padding:10px 13px;font-size:13px;outline:none}
.wp-page .wp-input:focus{border-color:var(--wpa-accent)}
.wp-page .wp-send{border:0;background:var(--wpa-accent);color:#fff;border-radius:12px;padding:0 18px;
  font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap}
.wp-page .wp-send[disabled]{opacity:.55;cursor:default}
.wp-page .wp-preview-body{padding:12px 14px 14px;gap:10px}
.wp-page .wp-stage{flex:1;min-height:0;overflow:auto;display:flex;justify-content:center;align-items:flex-start;
  background:var(--wpa-surface2);border:1px solid var(--wpa-line);border-radius:12px;padding:14px}
.wp-page .wp-paper{background:#fff;border-radius:4px;overflow:hidden;box-shadow:0 12px 34px rgba(0,0,0,.16);
  max-width:100%}
.wp-page .wp-paper img{display:block;width:100%;height:auto}
.wp-page .wp-empty{display:flex;flex-direction:column;align-items:center;gap:10px;margin:auto;color:var(--wpa-muted);
  font-size:12px;text-align:center;line-height:1.6;padding:20px}
.wp-page .wp-empty-icon{width:58px;height:58px;border-radius:16px;display:grid;place-items:center;
  color:var(--wpa-accent);border:1px dashed color-mix(in oklab, var(--wpa-accent) 40%, var(--wpa-line));
  background:color-mix(in oklab, var(--wpa-accent) 6%, var(--wpa-surface))}
.wp-page .wp-steplist{display:flex;gap:7px;overflow:auto;padding:2px 0 2px}
.wp-page .wp-step{display:inline-flex;align-items:center;gap:6px;white-space:nowrap;border:1px solid var(--wpa-line);
  background:var(--wpa-surface);border-radius:10px;padding:5px 9px;font-size:11.5px;color:var(--wpa-muted);
  cursor:pointer;font:inherit;font-size:11.5px;flex:0 0 auto;overflow:visible;text-overflow:clip;
  transition:border-color .15s, color .15s}
.wp-page .wp-step:hover{border-color:var(--wpa-accent)}
.wp-page .wp-step[data-active="1"]{border-color:var(--wpa-accent);color:var(--wpa-accent);
  background:color-mix(in oklab, var(--wpa-accent) 8%, var(--wpa-surface))}
.wp-page .wp-step-num{font-family:ui-monospace,Consolas,monospace;opacity:.65}
.wp-page .wp-pill-dot[data-busy="1"]{background:var(--dsw-alias-state-warn-primary)}
/* ---- DSH primitives skin: colors / font sizes / radii / shadows come from --dsw-* tokens ---- */
.wp-page{--wpa-line:var(--dsw-alias-border-l1);--wpa-strong:var(--dsw-alias-border-l2);
  --wpa-text:var(--dsw-alias-label-primary);--wpa-muted:var(--dsw-alias-label-secondary);
  --wpa-surface:var(--dsw-alias-bg-layer-1);--wpa-surface2:var(--dsw-alias-bg-layer-2);
  --wpa-accent:var(--dsw-alias-brand-primary);--wpa-on-accent:var(--dsw-alias-label-primary-inverted);
  background:var(--dsw-alias-bg-base);color:var(--wpa-text);
  font-family:var(--dsw-font-family);font-size:var(--dsw-font-s-14-font-size);
  line-height:var(--dsw-font-s-14-line-height)}
.wp-page .wp-page-inner{max-width:1120px;padding:32px 28px 40px;gap:20px}
.wp-page .wp-logo{width:auto;height:auto;border-radius:0;color:inherit;
  background:transparent;border:0;box-shadow:none}
.wp-page .wp-hero-title{font-size:var(--dsw-font-l-20-font-size);line-height:var(--dsw-font-l-20-line-height);
  font-weight:var(--dsw-font-l-20-font-weight);letter-spacing:0}
.wp-page .wp-hero-sub{font-size:var(--dsw-font-xs-13-font-size);line-height:var(--dsw-font-xs-13-line-height);
  color:var(--wpa-muted)}
.wp-page .wp-pill{font-size:var(--dsw-font-xxs-12-font-size);line-height:var(--dsw-font-xxs-12-line-height);
  color:var(--wpa-muted);background:var(--wpa-surface);border:1px solid var(--wpa-line);
  border-radius:var(--dsw-radius-panel);padding:4px 10px}
.wp-page .wp-row{background:var(--wpa-surface);border:1px solid var(--wpa-line);
  border-radius:var(--dsw-radius-lg);box-shadow:var(--dsw-shadow-lv1);padding:16px}
.wp-page .wp-row:hover{background:var(--dsw-alias-interactive-bg-hover);
  border-color:var(--wpa-strong);box-shadow:var(--dsw-shadow-lv2);transform:none}
.wp-page .wp-row-icon{width:40px;height:40px;border-radius:var(--dsw-radius-md);color:var(--wpa-text);
  background:var(--wpa-surface2);border:1px solid var(--wpa-line)}
.wp-page .wp-row-title{font-size:var(--dsw-font-s-strong-14-font-size);
  line-height:var(--dsw-font-s-strong-14-line-height);font-weight:var(--dsw-font-s-strong-14-font-weight)}
.wp-page .wp-row-desc{font-size:var(--dsw-font-xs-13-font-size);
  line-height:var(--dsw-font-xs-13-line-height);color:var(--wpa-muted)}
.wp-page .wp-row-meta{font-size:var(--dsw-font-xxxs-11-font-size);
  line-height:var(--dsw-font-xxxs-11-line-height);color:var(--wpa-muted);
  border:1px solid var(--wpa-line);border-radius:var(--dsw-radius-panel);padding:2px 8px}
.wp-page .wp-row-note{font-size:var(--dsw-font-xxs-12-font-size);line-height:var(--dsw-font-xxs-12-line-height);
  color:var(--wpa-muted)}
.wp-page .wp-panel{background:var(--wpa-surface);border:1px solid var(--wpa-line);
  border-radius:var(--dsw-radius-lg);box-shadow:var(--dsw-shadow-lv1)}
.wp-page .wp-panel-head{border-bottom:1px solid var(--wpa-line);padding:12px 16px}
.wp-page .wp-panel-title{font-size:var(--dsw-font-xs-strong-13-font-size);
  line-height:var(--dsw-font-xs-strong-13-line-height);font-weight:var(--dsw-font-xs-strong-13-font-weight);
  color:var(--wpa-text)}
.wp-page .wp-panel-meta{font-size:var(--dsw-font-xxxs-11-font-size);
  line-height:var(--dsw-font-xxxs-11-line-height);color:var(--wpa-muted)}
.wp-page .wp-drop{background:var(--wpa-surface2);border:1px dashed var(--wpa-strong);
  border-radius:var(--dsw-radius-md);padding:12px;min-height:68px;flex-direction:row;align-items:center;gap:10px}
.wp-page .wp-drop:hover{background:var(--dsw-alias-interactive-bg-hover);
  border-color:var(--dsw-alias-brand-primary)}
.wp-page .wp-drop-icon{width:32px;height:32px;border-radius:var(--dsw-radius-sm);
  color:var(--wpa-text);background:var(--wpa-surface);border:1px solid var(--wpa-line)}
.wp-page .wp-drop-label{font-size:var(--dsw-font-xs-strong-13-font-size);
  line-height:var(--dsw-font-xs-strong-13-line-height);font-weight:var(--dsw-font-xs-strong-13-font-weight)}
.wp-page .wp-drop-hint{font-size:var(--dsw-font-xxxs-11-font-size);
  line-height:var(--dsw-font-xxxs-11-line-height);color:var(--wpa-muted)}
.wp-page .wp-msg{font-size:var(--dsw-font-s-14-font-size);line-height:var(--dsw-font-s-14-line-height);
  border-radius:var(--dsw-radius-md);padding:8px 12px;max-width:92%}
.wp-page .wp-msg[data-role="agent"]{background:var(--wpa-surface2);border:1px solid var(--wpa-line);
  color:var(--wpa-text)}
.wp-page .wp-msg[data-role="user"]{background:var(--dsw-specific-bubble);color:var(--wpa-text);
  border:1px solid var(--dsw-alias-border-l2);box-shadow:none}
.wp-page .wp-chip{border:1px solid var(--wpa-line);background:var(--wpa-surface2);color:var(--wpa-text);
  border-radius:var(--dsw-radius-xs);padding:4px 10px;font-size:var(--dsw-font-xxs-12-font-size);
  line-height:var(--dsw-font-xxs-12-line-height)}
.wp-page .wp-chip:hover{background:var(--dsw-alias-interactive-bg-hover);
  border-color:var(--wpa-strong);color:var(--wpa-text)}
.wp-page .wp-input{background:var(--dsw-specific-input-major);border:1px solid var(--wpa-strong);
  border-radius:var(--dsw-radius-sm);padding:8px 12px;font-size:var(--dsw-font-s-14-font-size);
  line-height:var(--dsw-font-s-14-line-height);color:var(--wpa-text);outline:none}
.wp-page .wp-input::placeholder{color:var(--wpa-muted)}
.wp-page .wp-input:focus{border-color:var(--dsw-alias-brand-primary);
  box-shadow:0 0 0 var(--dsw-focus-ring-width) var(--dsw-alias-interactive-bg-hover-accent)}
.wp-page .wp-send{background:var(--dsw-alias-button-primary-fill);color:var(--wpa-on-accent);
  border-radius:var(--dsw-radius-sm);padding:0 16px;font-size:var(--dsw-font-xs-strong-13-font-size);
  font-weight:var(--dsw-font-xs-strong-13-font-weight)}
.wp-page .wp-send:hover:not([disabled]){background:var(--dsw-alias-button-primary-hover)}
.wp-page .wp-send[disabled]{background:var(--dsw-alias-button-primary-dimmed);opacity:1}
.wp-page .wp-stage{background:var(--dsw-alias-bg-document-preview);border:1px solid var(--wpa-line);
  border-radius:var(--dsw-radius-md);padding:12px}
.wp-page .wp-paper{background:var(--dsw-alias-bg-layer-1);border-radius:var(--dsw-radius-xs);
  box-shadow:var(--dsw-shadow-lv3)}
.wp-page .wp-empty{font-size:var(--dsw-font-xs-13-font-size);line-height:var(--dsw-font-xs-13-line-height);
  color:var(--wpa-muted)}
.wp-page .wp-empty-icon{width:56px;height:56px;border-radius:var(--dsw-radius-md);
  color:var(--wpa-text);background:var(--wpa-surface2);border:1px dashed var(--wpa-strong)}
.wp-page .wp-steplist{gap:6px;padding:2px 0}
.wp-page .wp-step{background:var(--wpa-surface2);border:1px solid var(--wpa-line);
  border-radius:var(--dsw-radius-xs);padding:4px 8px;color:var(--wpa-muted);
  font-size:var(--dsw-font-xxxs-11-font-size);line-height:var(--dsw-font-xxxs-11-line-height)}
.wp-page .wp-step:hover{border-color:var(--wpa-strong);color:var(--wpa-text)}
.wp-page .wp-step[data-active="1"]{border-color:var(--dsw-alias-brand-primary);
  background:var(--dsw-alias-interactive-bg-hover-accent);color:var(--wpa-text)}
.wp-page .wp-step-num{color:var(--wpa-muted)}
.wp-page .wp-error{font-size:var(--dsw-font-xs-13-font-size);
  line-height:var(--dsw-font-xs-13-line-height);color:var(--dsw-alias-state-error-primary)}
.wp-page .wp-btn{border:1px solid var(--wpa-strong);background:var(--dsw-alias-button-elevated-fill);
  color:var(--wpa-text);border-radius:var(--dsw-radius-sm);font-size:var(--dsw-font-xs-13-font-size);
  line-height:var(--dsw-font-xs-13-line-height)}
.wp-page .wp-btn:hover{border-color:var(--wpa-strong);background:var(--dsw-alias-interactive-bg-hover)}
.wp-page .wp-cta{width:100%;height:40px;border:0;border-radius:var(--dsw-radius-sm);
  background:var(--dsw-alias-button-primary-fill);color:var(--dsw-alias-label-primary-inverted);
  font-size:var(--dsw-font-s-strong-14-font-size);line-height:var(--dsw-font-s-strong-14-line-height);
  font-weight:var(--dsw-font-s-strong-14-font-weight);cursor:pointer}
.wp-page .wp-cta:hover:not([disabled]){background:var(--dsw-alias-button-primary-hover)}
.wp-page .wp-cta[disabled]{background:var(--dsw-alias-button-primary-dimmed);cursor:default}
.wp-page .wp-hint{font-size:var(--dsw-font-xxs-12-font-size);line-height:var(--dsw-font-xxs-12-line-height);
  color:var(--wpa-muted);text-align:center}
.wp-page .wp-progress{display:flex;flex-direction:column;gap:6px;padding:10px 12px;
  border:1px solid var(--wpa-line);border-radius:var(--dsw-radius-md);background:var(--wpa-surface2)}
.wp-page .wp-stage-row{display:flex;align-items:center;gap:8px;
  font-size:var(--dsw-font-xs-13-font-size);line-height:var(--dsw-font-xs-13-line-height);color:var(--wpa-muted)}
.wp-page .wp-stage-row[data-state="ok"]{color:var(--wpa-text)}
.wp-page .wp-stage-row[data-state="fail"]{color:var(--dsw-alias-state-error-primary)}
.wp-page .wp-stage-dot{width:8px;height:8px;border-radius:50%;flex:none;
  background:var(--dsw-alias-state-idle-primary)}
.wp-page .wp-stage-row[data-state="run"] .wp-stage-dot{background:var(--dsw-alias-state-warn-primary)}
.wp-page .wp-stage-row[data-state="ok"] .wp-stage-dot{background:var(--dsw-alias-state-success-primary)}
.wp-page .wp-stage-row[data-state="fail"] .wp-stage-dot{background:var(--dsw-alias-state-error-primary)}
.wp-page .wp-stage-meta{margin-left:auto;font-size:var(--dsw-font-xxxs-11-font-size);color:var(--wpa-muted)}
.wp-page .wp-dl-row{display:flex;gap:8px;flex-wrap:wrap;margin-top:2px}
.wp-page .wp-dl{display:inline-flex;align-items:center;justify-content:center;gap:6px;padding:7px 12px;
  border:1px solid var(--wpa-strong);border-radius:var(--dsw-radius-sm);
  background:var(--dsw-alias-button-elevated-fill);color:var(--wpa-text);text-decoration:none;
  font-size:var(--dsw-font-xs-strong-13-font-size);line-height:var(--dsw-font-xs-strong-13-line-height)}
.wp-page .wp-dl:hover{background:var(--dsw-alias-interactive-bg-hover)}
.wp-page .wp-dl-primary{background:var(--dsw-alias-button-primary-fill);
  color:var(--dsw-alias-label-primary-inverted);border-color:transparent}
.wp-page .wp-dl-primary:hover{background:var(--dsw-alias-button-primary-hover)}
.wp-page .wp-divider{height:1px;background:var(--wpa-line);margin:2px 0}
.wp-page .wp-subhead{font-size:var(--dsw-font-xxs-12-font-size);line-height:var(--dsw-font-xxs-12-line-height);
  color:var(--wpa-muted)}
.wp-page .wp-chat-compact{min-height:88px;max-height:150px;flex:none}
/* ---- accent: deepseek blue + amber, all from --dsw-static-* tokens ---- */
.wp-page{--wpa-blue:var(--dsw-static-deepseek-500);--wpa-blue-hover:var(--dsw-static-deepseek-400);
  --wpa-blue-soft:var(--dsw-static-deepseek-50);--wpa-blue-line:var(--dsw-static-deepseek-200);
  --wpa-blue-text:var(--dsw-static-deepseek-600);--wpa-white:var(--dsw-static-neutral-00)}
.wp-page .wp-row-icon{background:transparent;border:0;width:auto;height:auto;border-radius:0}
.wp-page .wp-row:hover{border-color:var(--wpa-blue-line)}
.wp-page .wp-cta{background:var(--wpa-blue);color:var(--wpa-white)}
.wp-page .wp-cta:hover:not([disabled]){background:var(--wpa-blue-hover)}
.wp-page .wp-cta[disabled]{background:var(--wpa-blue-line);color:var(--wpa-white)}
.wp-page .wp-dl-primary{background:var(--wpa-blue);color:var(--wpa-white);border-color:transparent}
.wp-page .wp-dl-primary:hover{background:var(--wpa-blue-hover)}
.wp-page .wp-drop-icon{background:var(--wpa-blue-soft);border-color:var(--wpa-blue-line);
  color:var(--wpa-blue-text)}
.wp-page .wp-drop:hover{border-color:var(--wpa-blue-line);background:var(--wpa-blue-soft)}
.wp-page .wp-msg[data-role="user"]{background:var(--wpa-blue-soft);border:1px solid var(--wpa-blue-line);
  color:var(--dsw-static-deepseek-900)}
.wp-page .wp-send{background:var(--wpa-blue-soft);color:var(--wpa-blue-text);
  border:1px solid var(--wpa-blue-line);border-radius:var(--dsw-radius-sm);
  font-size:var(--dsw-font-xs-strong-13-font-size);font-weight:var(--dsw-font-xs-strong-13-font-weight)}
.wp-page .wp-send:hover:not([disabled]){background:var(--wpa-blue-line)}
.wp-page .wp-send[disabled]{opacity:.6;background:var(--wpa-blue-soft);color:var(--wpa-blue-text)}
.wp-page .wp-step[data-active="1"]{border-color:var(--wpa-blue-line);background:var(--wpa-blue-soft);
  color:var(--wpa-blue-text)}
.wp-page .wp-empty-icon{background:var(--wpa-blue-soft);border-color:var(--wpa-blue-line);
  color:var(--wpa-blue-text)}
/* ---- frame: give the canvas and the cards separate surfaces so edges never disappear ---- */
.wp-page{background:var(--dsw-alias-bg-module-platform)}
.wp-page .wp-panel{border:1px solid color-mix(in oklab, var(--dsw-static-violet-600) 24%, transparent);
  box-shadow:var(--dsw-elevation-panel);background:var(--dsw-alias-bg-layer-1)}
.wp-page .wp-panel-head{border-bottom:1px solid color-mix(in oklab, var(--dsw-static-violet-600) 14%,
  transparent)}
.wp-page .wp-row{border:1px solid color-mix(in oklab, var(--dsw-static-violet-600) 22%, transparent);
  box-shadow:var(--dsw-elevation-soft)}
.wp-page .wp-row:hover{border-color:color-mix(in oklab, var(--dsw-static-violet-600) 46%, transparent)}
.wp-page .wp-drop{border:1px dashed color-mix(in oklab, var(--dsw-static-violet-600) 40%, transparent);
  background:var(--dsw-alias-bg-layer-1)}
.wp-page .wp-drop:hover{border-color:var(--dsw-static-deepseek-500);background:var(--wpa-blue-soft)}
.wp-page .wp-progress{border:1px solid color-mix(in oklab, var(--dsw-static-violet-600) 24%, transparent);
  background:var(--dsw-alias-bg-layer-1)}
.wp-page :focus-visible{outline:2px solid var(--dsw-static-deepseek-500);outline-offset:2px}
.wp-page .wp-drop:focus-visible{border-color:var(--dsw-static-deepseek-500)}
.wp-page .wp-cta:focus-visible,.wp-page .wp-send:focus-visible,.wp-page .wp-btn:focus-visible,
.wp-page .wp-dl:focus-visible{outline-offset:2px}
/* ---- premium band: dark gradient header + saturated accents ---- */
.wp-page{background:
  radial-gradient(1100px 420px at 8% -12%,
    color-mix(in oklab, var(--dsw-static-violet-600) 20%, transparent), transparent 60%),
  radial-gradient(900px 380px at 102% -6%,
    color-mix(in oklab, var(--dsw-static-deepseek-500) 18%, transparent), transparent 58%),
  var(--dsw-alias-bg-module-platform)}
.wp-page .wp-hero-band{position:relative;overflow:hidden;border-radius:var(--dsw-radius-lg);
  padding:20px 22px;box-shadow:var(--dsw-shadow-lv3);
  background:linear-gradient(135deg,
    var(--dsw-static-neutral-bluish-1000) 0%,
    var(--dsw-static-deepseek-800) 52%,
    var(--dsw-static-violet-600) 100%)}
.wp-page .wp-hero-band::after{content:'';position:absolute;inset:0;pointer-events:none;
  background:radial-gradient(420px 170px at 86% 8%,
    color-mix(in oklab, var(--dsw-static-violet-400) 42%, transparent), transparent 72%)}
.wp-page .wp-hero-band .wp-hero{position:relative;z-index:1}
.wp-page .wp-hero-band .wp-hero-title{color:var(--dsw-static-neutral-00)}
.wp-page .wp-hero-band .wp-hero-sub{color:color-mix(in oklab, var(--dsw-static-neutral-00) 78%, transparent)}
.wp-page .wp-hero-band .wp-logo{box-shadow:0 0 0 1px color-mix(in oklab, var(--dsw-static-neutral-00) 24%, transparent)}
.wp-page .wp-hero-band .wp-pill{color:var(--dsw-static-neutral-00);
  background:color-mix(in oklab, var(--dsw-static-neutral-00) 14%, transparent);
  border-color:color-mix(in oklab, var(--dsw-static-neutral-00) 28%, transparent)}
.wp-page .wp-hero-band .wp-btn{color:var(--dsw-static-neutral-00);
  background:color-mix(in oklab, var(--dsw-static-neutral-00) 14%, transparent);
  border-color:color-mix(in oklab, var(--dsw-static-neutral-00) 30%, transparent)}
.wp-page .wp-hero-band .wp-btn:hover{
  background:color-mix(in oklab, var(--dsw-static-neutral-00) 24%, transparent)}
.wp-page .wp-cta{background:linear-gradient(135deg,
  var(--dsw-static-deepseek-500), var(--dsw-static-violet-600));
  box-shadow:0 6px 18px color-mix(in oklab, var(--dsw-static-violet-600) 34%, transparent)}
.wp-page .wp-cta:hover:not([disabled]){background:linear-gradient(135deg,
  color-mix(in oklab, var(--dsw-static-deepseek-500) 86%, var(--dsw-static-neutral-00)),
  color-mix(in oklab, var(--dsw-static-violet-600) 86%, var(--dsw-static-neutral-00)))}
.wp-page .wp-dl-primary{background:linear-gradient(135deg,
  var(--dsw-static-deepseek-500), var(--dsw-static-violet-600));
  box-shadow:0 4px 12px color-mix(in oklab, var(--dsw-static-violet-600) 26%, transparent)}
.wp-page .wp-progress{background:color-mix(in oklab, var(--dsw-static-violet-600) 5%,
    var(--dsw-alias-bg-layer-1));
  border-color:color-mix(in oklab, var(--dsw-static-violet-600) 16%, var(--dsw-alias-border-l1))}
.wp-page .wp-panel{border-top:2px solid color-mix(in oklab, var(--dsw-static-violet-600) 26%,
    transparent)}
/* ---- tinted surfaces: panels get a light brand wash instead of white ---- */
.wp-page .wp-panel{background:color-mix(in oklab, var(--dsw-static-violet-600) 6%,
  var(--dsw-alias-bg-layer-1))}
.wp-page .wp-workspace > .wp-panel:nth-child(1){background:color-mix(in oklab,
  var(--dsw-static-deepseek-500) 8%, var(--dsw-alias-bg-layer-1));
  border-color:color-mix(in oklab, var(--dsw-static-deepseek-500) 30%, transparent)}
.wp-page .wp-workspace > .wp-panel:nth-child(2){background:color-mix(in oklab,
  var(--dsw-static-violet-600) 9%, var(--dsw-alias-bg-layer-1));
  border-color:color-mix(in oklab, var(--dsw-static-violet-600) 30%, transparent)}
.wp-page .wp-panel-head{background:transparent}
.wp-page .wp-row{background:color-mix(in oklab, var(--dsw-static-violet-600) 3%,
  var(--dsw-alias-bg-layer-1))}
.wp-page .wp-drop{background:var(--dsw-alias-bg-layer-1)}
.wp-page .wp-stage{background:color-mix(in oklab, var(--dsw-static-violet-600) 5%,
  var(--dsw-alias-bg-document-preview))}
.wp-page .wp-progress{background:color-mix(in oklab, var(--dsw-static-violet-600) 5%,
  var(--dsw-alias-bg-layer-1))}
/* ---- typography: display family for headings, tuned scale ---- */
.wp-page{--wp-font-stack:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Hiragino Sans GB","Microsoft YaHei","Helvetica Neue",Arial,"Noto Sans",sans-serif;font-family:var(--wp-font-stack);-webkit-font-smoothing:auto;-moz-osx-font-smoothing:auto}
.wp-page .wp-hero-band{padding:22px 24px}
.wp-page .wp-hero-title{font-family:var(--wp-font-stack);
  font-size:var(--dsw-font-xl-24-font-size);line-height:var(--dsw-font-xl-24-line-height);
  font-weight:var(--dsw-font-xl-24-font-weight);letter-spacing:0}
.wp-page .wp-hero-sub{font-size:var(--dsw-font-xs-13-font-size);
  line-height:var(--dsw-font-xs-13-line-height);letter-spacing:0}
.wp-page .wp-panel-title{font-family:var(--wp-font-stack);
  font-size:var(--dsw-font-xs-strong-13-font-size);line-height:var(--dsw-font-xs-strong-13-line-height);
  font-weight:600;letter-spacing:0}
.wp-page .wp-row-title{font-family:var(--wp-font-stack);
  font-size:var(--dsw-font-m-18-font-size);line-height:1.35;font-weight:600;letter-spacing:0}
.wp-page .wp-row-desc{font-size:var(--dsw-font-s-14-font-size);
  line-height:var(--dsw-font-s-14-line-height)}
.wp-page .wp-cta{font-size:var(--dsw-font-s-strong-14-font-size);
  line-height:var(--dsw-font-s-strong-14-line-height);font-weight:600;letter-spacing:0}
.wp-page .wp-msg{font-size:var(--dsw-font-s-14-font-size);line-height:var(--dsw-font-s-14-line-height)}
.wp-page .wp-step{font-family:var(--ds-font-family-code, ui-monospace, Consolas, monospace);
  font-size:var(--dsw-font-xxxs-11-font-size)}
.wp-page .wp-pill{letter-spacing:0}
.wp-page .wp-empty{font-size:var(--dsw-font-s-14-font-size);line-height:var(--dsw-font-s-14-line-height)}
.wp-page .wp-drop-label{letter-spacing:0}
.wp-page .wp-loaded{display:flex;flex-direction:column;gap:6px;padding:10px 12px;
  border:1px solid color-mix(in oklab, var(--dsw-static-deepseek-500) 22%, transparent);
  border-radius:var(--dsw-radius-md);
  background:color-mix(in oklab, var(--dsw-static-deepseek-500) 5%, var(--dsw-alias-bg-layer-1))}
.wp-page .wp-loaded-head{font-size:var(--dsw-font-xxxs-11-font-size);
  line-height:var(--dsw-font-xxxs-11-line-height);color:var(--wpa-muted)}
.wp-page .wp-loaded-row{display:flex;align-items:center;gap:8px;
  font-size:var(--dsw-font-xs-13-font-size);line-height:var(--dsw-font-xs-13-line-height)}
.wp-page .wp-loaded-tag{flex:none;padding:1px 7px;border-radius:var(--dsw-radius-xs);
  background:var(--wpa-blue-soft);color:var(--wpa-blue-text);font-size:var(--dsw-font-xxxs-11-font-size)}
.wp-page .wp-loaded-name{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
  color:var(--wpa-text)}
.wp-page .wp-loaded-state{flex:none;font-size:var(--dsw-font-xxxs-11-font-size);color:var(--wpa-muted)}
.wp-page .wp-loaded-row[data-uploaded="1"] .wp-loaded-state{color:var(--dsw-alias-state-success-primary)}
.wp-page .wp-loaded-x{flex:none;width:20px;height:20px;border-radius:var(--dsw-radius-xs);
  border:1px solid var(--wpa-line);background:transparent;color:var(--wpa-muted);cursor:pointer;
  font-size:12px;line-height:1}
.wp-page .wp-loaded-x:hover{color:var(--dsw-alias-state-error-primary);
  border-color:var(--dsw-alias-state-error-primary)}
.wp-page .wp-drop-wrap{display:flex;flex-direction:column;gap:6px;min-width:0}
.wp-page .wp-drop{width:100%;box-sizing:border-box}
.wp-page .wp-drop[data-loaded="1"]{border-style:solid;
  border-color:color-mix(in oklab, var(--dsw-static-deepseek-500) 42%, transparent);
  background:color-mix(in oklab, var(--dsw-static-deepseek-500) 6%, var(--dsw-alias-bg-layer-1))}
.wp-page .wp-drop-clear{flex:none;width:20px;height:20px;border-radius:var(--dsw-radius-xs);
  border:1px solid var(--wpa-line);background:transparent;color:var(--wpa-muted);cursor:pointer;
  font-size:12px;line-height:1}
.wp-page .wp-drop-clear:hover{color:var(--dsw-alias-state-error-primary);
  border-color:var(--dsw-alias-state-error-primary)}
.wp-page .wp-drop-track{height:4px;border-radius:999px;overflow:hidden;
  background:color-mix(in oklab, var(--dsw-static-violet-600) 12%, transparent)}
.wp-page .wp-drop-track-fill{display:block;height:100%;width:0;
  background:linear-gradient(90deg, var(--dsw-static-deepseek-500), var(--dsw-static-violet-600));
  transition:width .2s ease}
.wp-page .wp-drop-track[data-loading="1"]{background:color-mix(in oklab, var(--dsw-static-violet-600) 20%, transparent)}
.wp-page .wp-drop-error{font-size:var(--dsw-font-xxxs-11-font-size);line-height:1.45;
  color:var(--dsw-alias-state-error-primary);word-break:break-word}
.wp-page .wp-stage img{cursor:zoom-in}
.wp-page .wp-lightbox{position:fixed;inset:0;z-index:1200;display:flex;align-items:center;
  justify-content:center;padding:28px;background:rgba(9,10,14,.78)}
.wp-page .wp-lightbox img{max-width:94vw;max-height:92vh;width:auto;height:auto;background:#fff;
  border-radius:6px;box-shadow:0 24px 60px rgba(0,0,0,.45);cursor:zoom-out}
.wp-page .wp-lightbox-close{position:absolute;top:18px;right:22px;width:32px;height:32px;
  border-radius:50%;border:1px solid rgba(255,255,255,.35);background:rgba(255,255,255,.12);
  color:#fff;font-size:18px;line-height:1;cursor:pointer}
.wp-page .wp-log{border:1px solid color-mix(in oklab, var(--dsw-static-violet-600) 18%, transparent);
  border-radius:var(--dsw-radius-md);background:color-mix(in oklab, var(--dsw-alias-bg-base) 70%,
  var(--dsw-alias-bg-layer-1));display:flex;flex-direction:column;min-height:0;order:-1}
.wp-page .wp-log-head{display:flex;align-items:center;justify-content:space-between;padding:6px 10px;
  border-bottom:1px solid color-mix(in oklab, var(--dsw-static-violet-600) 12%, transparent);
  font-size:var(--dsw-font-xxs-12-font-size);color:var(--wpa-text);font-weight:600}
.wp-page .wp-log-count{color:var(--wpa-muted);font-weight:400;font-size:var(--dsw-font-xxxs-11-font-size)}
.wp-page .wp-log-body{max-height:150px;overflow:auto;padding:6px 10px;display:flex;flex-direction:column;gap:3px}
.wp-page .wp-log-row{display:flex;gap:8px;align-items:baseline;
  font-family:var(--ds-font-family-code, ui-monospace, Consolas, monospace);
  font-size:11px;line-height:1.5;color:var(--wpa-muted)}
.wp-page .wp-log-time{color:var(--wpa-muted);opacity:.8;flex:none}
.wp-page .wp-log-phase{flex:none;color:var(--dsw-static-deepseek-500)}
.wp-page .wp-log-msg{flex:1;min-width:0;color:var(--wpa-text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.wp-page .wp-log-ms{flex:none;color:var(--wpa-muted);opacity:.8}
.wp-page .wp-log-row[data-level="error"] .wp-log-msg{color:var(--dsw-alias-state-error-primary)}
.wp-page .wp-log-row[data-level="warn"] .wp-log-msg{color:var(--dsw-alias-state-warn-primary)}
.wp-page .wp-log-empty{padding:10px;color:var(--wpa-muted);font-size:var(--dsw-font-xxs-12-font-size)}
.wp-page .wp-word{flex:1;min-height:0;display:flex;flex-direction:column;align-items:center;
  justify-content:center;gap:10px;padding:26px;text-align:center;
  background:var(--wpa-surface2);border:1px solid var(--wpa-line);border-radius:12px}
.wp-page .wp-word-icon{display:flex}
.wp-page .wp-word-title{font-size:var(--dsw-font-m-18-font-size);font-weight:600;color:var(--wpa-text)}
.wp-page .wp-word-sub{max-width:340px;font-size:var(--dsw-font-xs-13-font-size);line-height:1.6;
  color:var(--wpa-muted)}
.wp-page .wp-word-open{appearance:none;border:0;border-radius:12px;padding:10px 18px;cursor:pointer;
  font-size:var(--dsw-font-s-14-font-size);font-weight:600;color:var(--dsw-static-neutral-00);
  background:linear-gradient(135deg, var(--dsw-static-deepseek-500), var(--dsw-static-violet-600))}
.wp-page .wp-word-open:hover{filter:brightness(1.06)}
.wp-page .wp-word-map{appearance:none;border:1px solid var(--wpa-line);background:transparent;
  color:var(--wpa-text);border-radius:12px;padding:9px 16px;cursor:pointer;font-size:13px}
.wp-page .wp-word-map:hover{border-color:var(--wpa-accent);color:var(--wpa-accent)}
.wp-page .wp-map-modal{width:min(780px,92vw);max-height:86vh;display:flex;flex-direction:column;
  background:var(--wpa-surface);border-radius:14px;overflow:hidden;box-shadow:0 24px 60px rgba(0,0,0,.4)}
.wp-page .wp-map-head{display:flex;align-items:center;justify-content:space-between;
  padding:14px 18px;border-bottom:1px solid var(--wpa-line);font-weight:600}
.wp-page .wp-map-body{padding:14px 18px;overflow:auto;display:flex;flex-direction:column;gap:10px}
.wp-page .wp-map-notes{font-size:13px;color:var(--wpa-muted);line-height:1.6}
.wp-page .wp-map-row{display:grid;grid-template-columns:1.2fr auto 1.2fr auto;gap:10px;align-items:center;
  padding:5px 0;border-bottom:1px dashed var(--wpa-line);
  font-family:var(--ds-font-family-code, ui-monospace, Consolas, monospace);font-size:12px}
.wp-page .wp-map-arrow{color:var(--wpa-muted)}
.wp-page .wp-map-a{color:var(--dsw-static-deepseek-500)}
.wp-page .wp-map-issue{font-size:12px;line-height:1.6;color:var(--wpa-muted)}
.wp-page .wp-map-issue[data-level="error"]{color:var(--dsw-alias-state-error-primary)}
.wp-page .wp-map-issue[data-level="warn"]{color:var(--dsw-alias-state-warn-primary)}
.wp-page .wp-map-raw summary{cursor:pointer;font-size:12px;color:var(--wpa-muted)}
.wp-page .wp-map-raw pre{max-height:220px;overflow:auto;background:var(--wpa-surface2);
  border-radius:8px;padding:10px;font-size:11.5px;white-space:pre-wrap;word-break:break-all}
.wp-page .wp-check-issues{margin-top:8px;padding:8px 10px;border-radius:var(--dsw-radius-md);
  background:color-mix(in oklab, var(--dsw-alias-state-error-primary) 8%, transparent);
  display:flex;flex-direction:column;gap:4px}
.wp-page .wp-check-issue{font-size:11.5px;line-height:1.5;color:var(--dsw-alias-state-error-primary);
  word-break:break-word}
@media (max-width:1200px){
  .wp-page .wp-workspace{grid-template-columns:minmax(0,1fr)}
  .wp-page .wp-page-inner{padding:20px 16px 28px}
}
`;

    function toBase64(buffer) {
      const bytes = new Uint8Array(buffer);
      let binary = '';
      const chunk = 0x8000;
      for (let index = 0; index < bytes.length; index += chunk) {
        binary += String.fromCharCode.apply(null, bytes.subarray(index, index + chunk));
      }
      return btoa(binary);
    }

    async function call(route, body) {
      const response = await fetch(API + route, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Word-Token': token() },
        body: JSON.stringify(body),
      });
      const data = await response.json().catch(() => ({ ok: false, error: '响应不是 JSON' }));
      if (!response.ok || data.ok === false) {
        throw new Error(data.error || `HTTP ${response.status}`);
      }
      return data;
    }

    async function streamAgentRun(body, onEvent) {
      const response = await fetch(API + '/agent/run-stream', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Word-Token': token() },
        body: JSON.stringify(body),
      });
      if (!response.ok || !response.body) throw new Error(`Agent 请求失败（HTTP ${response.status}）`);
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        let index = buffer.indexOf('\n');
        while (index >= 0) {
          const line = buffer.slice(0, index).trim();
          buffer = buffer.slice(index + 1);
          index = buffer.indexOf('\n');
          if (!line) continue;
          let event = null;
          try {
            event = JSON.parse(line);
          } catch {
            continue;
          }
          if (event.type === 'error') throw new Error(event.error);
          if (onEvent) onEvent(event);
        }
      }
    }

    async function streamConvert(body, onEvent) {
      const response = await fetch(API + '/convert-stream', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Word-Token': token() },
        body: JSON.stringify(body),
      });
      if (!response.ok || !response.body) {
        throw new Error(`转换请求失败（HTTP ${response.status}）`);
      }
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      let result = null;
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        let index = buffer.indexOf('\n');
        while (index >= 0) {
          const line = buffer.slice(0, index).trim();
          buffer = buffer.slice(index + 1);
          index = buffer.indexOf('\n');
          if (!line) continue;
          let event = null;
          try {
            event = JSON.parse(line);
          } catch {
            continue;
          }
          if (event.type === 'done') result = event.result;
          else if (event.type === 'error') throw new Error(event.error);
          else if (onEvent) onEvent(event);
        }
      }
      if (!result) throw new Error('转换流意外结束');
      return result;
    }

    async function uploadFile(file, kind, onProgress) {
      const data = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
        reader.onerror = () => reject(new Error('读取文件失败'));
        reader.readAsDataURL(file);
      });
      return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open('POST', API + '/upload');
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.setRequestHeader('X-Word-Token', token());
        xhr.upload.onprogress = (event) => {
          if (event.lengthComputable && onProgress) {
            onProgress(Math.max(1, Math.min(99, Math.round((event.loaded / event.total) * 100))));
          }
        };
        xhr.onload = () => {
          let payload = null;
          try {
            payload = JSON.parse(xhr.responseText);
          } catch {
            payload = null;
          }
          if (xhr.status >= 200 && xhr.status < 300 && payload && payload.ok) resolve(payload);
          else reject(new Error((payload && payload.error) || `上传失败（HTTP ${xhr.status}）`));
        };
        xhr.onerror = () => reject(new Error('上传失败：网络错误'));
        xhr.send(JSON.stringify({ kind, name: file.name, data }));
      });
    }

    let dropSeq = 0;
    const dropTargets = new Map();

    function DropZone({ label, hint, accept, kind, onChange, onError, value }) {
      const input = React.useRef(null);
      const keyRef = React.useRef(null);
      if (keyRef.current === null) keyRef.current = `wp-drop-${++dropSeq}`;
      const [over, setOver] = React.useState(false);
      const [progress, setProgress] = React.useState(0);
      const [uploadError, setUploadError] = React.useState('');
      const loaded = Boolean(value && value.path);
      const loading = Boolean(value && value.loading);

      async function take(file) {
        if (!file) return;
        setUploadError('');
        setProgress(1);
        onChange({ name: file.name, path: '', loading: true });
        try {
          const saved = await uploadFile(file, kind, setProgress);
          onChange({ name: saved.name, path: saved.path, loading: false });
        } catch (error) {
          setProgress(0);
          onChange(null);
          const message = String(error.message || error);
          setUploadError(message);
          onError(message);
        }
      }

      React.useEffect(() => {
        dropTargets.set(keyRef.current, take);
        return () => dropTargets.delete(keyRef.current);
      });

      return h(
        'div',
        { className: 'wp-drop-wrap' },
        h(
          'div',
          {
            className: 'wp-drop',
            'data-kind': kind,
            'data-drop-key': keyRef.current,
            'data-over': over ? '1' : '0',
            'data-loaded': loaded ? '1' : '0',
            role: 'button',
            tabIndex: 0,
            onClick: () => input.current && input.current.click(),
            onKeyDown: (event) => {
              if (event.key === 'Enter' || event.key === ' ') input.current && input.current.click();
            },
            onDragOver: (event) => {
              event.preventDefault();
              setOver(true);
            },
            onDragLeave: () => setOver(false),
            onDrop: (event) => {
              event.preventDefault();
              setOver(false);
              take(event.dataTransfer.files && event.dataTransfer.files[0]);
            },
          },
          h(
            'span',
            { className: 'wp-drop-icon', 'aria-hidden': true },
            loaded ? '✓' : '＋',
          ),
          h(
            'span',
            { className: 'wp-drop-text' },
            h('span', { className: 'wp-drop-label' }, loaded ? '已加载文档' : label),
            h(
              'span',
              { className: 'wp-drop-hint' },
              loading ? `上传中 ${progress}%` : loaded ? value.name : hint,
            ),
          ),
          loaded
            ? h(
                'button',
                {
                  className: 'wp-drop-clear',
                  title: '清除，改用示例文件',
                  onClick: (event) => {
                    event.stopPropagation();
                    setProgress(0);
                    onChange(null);
                  },
                },
                '×',
              )
            : null,
          h('input', {
            ref: input,
            type: 'file',
            accept,
            style: { display: 'none' },
            onClick: (event) => event.stopPropagation(),
            onChange: (event) => {
              take(event.target.files && event.target.files[0]);
              event.target.value = '';
            },
          }),
        ),
        loading
          ? h(
              'div',
              { className: 'wp-drop-track', 'data-loading': '1' },
              h('span', { className: 'wp-drop-track-fill', style: { width: `${progress}%` } }),
            )
          : null,
        uploadError ? h('div', { className: 'wp-drop-error' }, uploadError) : null,
      );
    }

    function Panel() {
      const [files, setFiles] = React.useState({ template: null, source: null });
      const [state, setState] = React.useState({ status: 'idle', result: null, error: '', busy: '' });
      const [steps, setSteps] = React.useState([]);
      const [selected, setSelected] = React.useState(null);
      const [previewError, setPreviewError] = React.useState('');
      const follow = React.useRef(true);
      const alive = React.useRef(true);

      React.useEffect(() => {
        alive.current = true;
        fetch(API + '/state', { headers: { 'X-Word-Token': token() } })
          .then((response) => response.json())
          .then((data) => {
            if (alive.current && data && data.ok) setState((current) => ({ ...current, defaults: data.defaults }));
          })
          .catch(() => {});
        return () => {
          alive.current = false;
        };
      }, []);

      async function pullProgress() {
        try {
          const response = await fetch(API + '/progress', { headers: { 'X-Word-Token': token() } });
          const data = await response.json();
          if (!data || !data.ok || !Array.isArray(data.steps) || !data.steps.length) return;
          setSteps(data.steps.slice());
          if (follow.current) {
            const last = data.steps[data.steps.length - 1];
            if (last && last.step) setSelected(last.step);
          }
        } catch {
          /* the next tick retries */
        }
      }

      async function run(op) {
        setState((current) => ({ ...current, busy: op, error: '', status: 'run' }));
        setSteps([]);
        setSelected(null);
        setPreviewError('');
        follow.current = true;
        const polling = op === 'fill' ? setInterval(pullProgress, 350) : null;
        try {
          const result = await call('/op', {
            op,
            template: files.template && files.template.path,
            source: files.source && files.source.path,
          });
          if (result.steps && result.steps.length) {
            setSteps(result.steps.slice());
            const last = result.steps[result.steps.length - 1];
            if (last && last.step) setSelected(last.step);
          }
          setState((current) => ({
            ...current,
            busy: '',
            status: result.ok ? 'ok' : 'bad',
            result,
          }));
          if (op === 'fill' && result.steps && result.steps.length > 1) {
            replay(result.steps);
          }
        } catch (error) {
          setState((current) => ({
            ...current,
            busy: '',
            status: 'bad',
            error: String(error.message || error),
          }));
        } finally {
          if (polling) {
            clearInterval(polling);
            await pullProgress();
          }
        }
      }

      async function replay(list) {
        const items = Array.isArray(list) ? list : steps;
        follow.current = false;
        for (const item of items) {
          if (!item.step || !alive.current) continue;
          setSelected(item.step);
          setPreviewError('');
          await new Promise((resolve) => setTimeout(resolve, 900));
        }
      }

      const busy = state.busy;
      const statusText =
        state.status === 'run'
          ? `正在执行 ${busy}…`
          : state.status === 'ok'
            ? '执行成功'
            : state.status === 'bad'
              ? '执行失败'
              : '就绪';

      return h(
        'div',
        { className: 'wp-root' },
        h('style', null, CSS),
        h(
          'div',
          { className: 'wp-head' },
          h('span', { className: 'wp-title' }, 'Word 流水线'),
          h('span', { className: 'wp-badge' }, '本地'),
        ),
        h(
          'div',
          { className: 'wp-status' },
          h('span', { className: 'wp-dot', 'data-state': state.status }),
          h('span', null, statusText),
        ),
        h(
          'div',
          { className: 'wp-grid' },
          h(DropZone, {
            label: '资料 DOCX',
            hint: state.defaults ? '使用示例资料，或拖入 .docx' : '拖入 .docx',
            accept: '.docx',
            kind: 'source',
            value: files.source,
            onChange: (value) => setFiles((current) => ({ ...current, source: value })),
            onError: (message) => setState((current) => ({ ...current, error: message, status: 'bad' })),
          }),
          h(DropZone, {
            label: '模板 DOCX',
            hint: state.defaults ? '使用示例模板，或拖入 .docx' : '拖入 .docx',
            accept: '.docx',
            kind: 'template',
            value: files.template,
            onChange: (value) => setFiles((current) => ({ ...current, template: value })),
            onError: (message) => setState((current) => ({ ...current, error: message, status: 'bad' })),
          }),
        ),
        h(
          'div',
          { className: 'wp-actions' },
          ['fill', 'check', 'apply', 'render'].map((op) =>
            h(
              'button',
              {
                key: op,
                className: 'wp-btn',
                'data-busy': busy === op ? '1' : '0',
                disabled: Boolean(busy),
                onClick: () => run(op),
              },
              { fill: '填充', check: '检查', apply: '合批', render: '渲染' }[op],
            ),
          ),
        ),
        steps.length
          ? h(
              'div',
              { className: 'wp-section' },
              h(
                'div',
                { className: 'wp-section-title' },
                h('span', null, `修改过程 · ${steps.length} 步`),
                h(
                  'button',
                  {
                    className: 'wp-btn',
                    style: { padding: '1px 8px', fontSize: '11px' },
                    disabled: Boolean(busy),
                    onClick: replay,
                  },
                  '回放',
                ),
              ),
              h(
                'div',
                { className: 'wp-steps' },
                ...steps.map((item) =>
                  h(
                    'button',
                    {
                      key: `${item.op}:${item.step}`,
                      className: 'wp-step',
                      'data-active': selected === item.step ? '1' : '0',
                      onClick: () => {
                        follow.current = false;
                        setSelected(item.step);
                        setPreviewError('');
                      },
                    },
                    `${String(item.op).padStart(2, '0')} ${item.kind} · ${item.slot}`,
                  ),
                ),
              ),
              h(
                'div',
                { className: 'wp-stage' },
                selected
                  ? h('img', {
                      src: `${API}/preview?path=${encodeURIComponent(selected)}&t=${encodeURIComponent(token())}`,
                      alt: '当前步骤的 Word 渲染',
                      onError: () => setPreviewError('这一步的预览渲染失败（Word 可能正忙）'),
                    })
                  : h('div', { className: 'wp-stage-note' }, '运行填充后，这里按步骤显示文档'),
              ),
              previewError ? h('div', { className: 'wp-error' }, previewError) : null,
            )
          : null,
        state.error ? h('div', { className: 'wp-error' }, state.error) : null,
        state.result
          ? h(
              'div',
              { className: 'wp-files' },
              h('div', { className: 'wp-cmd' }, state.result.command || ''),
              h(
                'pre',
                { className: 'wp-log' },
                [state.result.stdout, state.result.stderr].filter(Boolean).join('\n') || '(无输出)',
              ),
              state.result.artifacts && state.result.artifacts.length
                ? h(
                    'div',
                    null,
                    '产物：',
                    ...state.result.artifacts.map((item) =>
                      h('div', { key: item, className: 'wp-path' }, item),
                    ),
                  )
                : null,
            )
          : null,
        h(
          'div',
          { className: 'wp-files' },
          h('span', null, '填充使用自动槽位映射（资料段落 → 模板槽位，按顺序）；合批使用示例修正队列。'),
        ),
      );
    }

    const overlay = { open: false, listeners: new Set() };

    function setOverlay(open) {
      overlay.open = open;
      overlay.listeners.forEach((listener) => listener(open));
    }

    function useOverlay() {
      const [open, setOpen] = React.useState(overlay.open);
      React.useEffect(() => {
        const listener = (value) => setOpen(value);
        overlay.listeners.add(listener);
        return () => overlay.listeners.delete(listener);
      }, []);
      return open;
    }

    function OverlayHost() {
      const open = useOverlay();
      if (!open) return null;
      return h(
        'div',
        {
          style: {
            position: 'fixed',
            top: '16px',
            right: '16px',
            bottom: '16px',
            width: '380px',
            zIndex: 60,
            display: 'flex',
            flexDirection: 'column',
            borderRadius: '14px',
            border: '1px solid var(--dsw-alias-border-l2)',
            background: 'var(--dsw-alias-bg-overlay)',
            boxShadow: '0 18px 48px rgba(0,0,0,.28)',
            overflow: 'hidden',
          },
        },
        h(
          'div',
          { style: { display: 'flex', justifyContent: 'flex-end', padding: '6px 8px 0' } },
          h(
            'button',
            {
              className: 'wp-btn',
              style: { padding: '2px 10px', fontSize: '12px' },
              onClick: () => setOverlay(false),
            },
            '关闭',
          ),
        ),
        h('div', { style: { flex: 1, minHeight: 0 } }, h(Panel)),
      );
    }

    function AgentIcon({ size = 20, active = false }) {
      return h(
        'svg',
        {
          width: size,
          height: size,
          viewBox: '0 0 24 24',
          fill: 'none',
          'aria-hidden': true,
          style: { display: 'block', flex: 'none' },
        },
        h('rect', {
          x: 0,
          y: 0,
          width: 24,
          height: 24,
          rx: 6,
          fill: active ? 'var(--dsw-static-deepseek-500)' : 'var(--dsw-static-deepseek-500)',
        }),
        h('path', {
          d: 'M7.1 5.6h5.5l3.7 3.7v9.1H7.1z',
          fill: 'var(--dsw-static-neutral-00)',
        }),
        h('path', {
          d: 'M12.6 5.6v3.7h3.7',
          stroke: 'var(--dsw-static-deepseek-500)',
          strokeWidth: 1.4,
          fill: 'none',
          strokeLinejoin: 'round',
        }),
        h('path', {
          d: 'M9.3 12.5h4.4M9.3 15h2.9',
          stroke: 'var(--dsw-static-deepseek-300)',
          strokeWidth: 1.2,
          strokeLinecap: 'round',
        }),
        h('circle', {
          cx: 17.2,
          cy: 17.2,
          r: 3.2,
          fill: 'var(--dsw-static-amber-400)',
        }),
        h('path', {
          d: 'M16 17.2h2.4M17.3 15.9l1.3 1.3-1.3 1.3',
          stroke: 'var(--dsw-static-neutral-00)',
          strokeWidth: 1.1,
          strokeLinecap: 'round',
          strokeLinejoin: 'round',
        }),
      );
    }

    function ManagerIcon({ size = 20 }) {
      return h(
        'svg',
        {
          width: size,
          height: size,
          viewBox: '0 0 24 24',
          fill: 'none',
          'aria-hidden': true,
          style: { display: 'block', flex: 'none' },
        },
        h('rect', {
          x: 0,
          y: 0,
          width: 24,
          height: 24,
          rx: 6,
          fill: 'var(--dsw-static-violet-600)',
        }),
        h('rect', {
          x: 5.4,
          y: 5.4,
          width: 5.6,
          height: 5.6,
          rx: 1.7,
          fill: 'var(--dsw-static-neutral-00)',
        }),
        h('rect', {
          x: 13,
          y: 5.4,
          width: 5.6,
          height: 5.6,
          rx: 1.7,
          fill: 'var(--dsw-static-neutral-00)',
          opacity: 0.7,
        }),
        h('rect', {
          x: 5.4,
          y: 13,
          width: 5.6,
          height: 5.6,
          rx: 1.7,
          fill: 'var(--dsw-static-neutral-00)',
          opacity: 0.7,
        }),
        h('rect', {
          x: 13,
          y: 13,
          width: 5.6,
          height: 5.6,
          rx: 1.7,
          fill: 'var(--dsw-static-neutral-00)',
        }),
      );
    }

    function AgentPanelPage() {
      const [view, setView] = React.useState('grid');
      const [files, setFiles] = React.useState({ template: null, source: null });
      const [messages, setMessages] = React.useState([
        {
          role: 'agent',
          text: '点「开始转换」即可按 填充 → 检查 → 渲染 自动生成成稿；哪里不合适，再在这里告诉我。',
        },
      ]);
      const [input, setInput] = React.useState('');
      const [busy, setBusy] = React.useState(false);
      const [steps, setSteps] = React.useState([]);
      const [preview, setPreview] = React.useState(null);
      const [error, setError] = React.useState('');
      const [outcome, setOutcome] = React.useState(null);
      const [liveCount, setLiveCount] = React.useState(0);
      const [stageText, setStageText] = React.useState('');
      const [logs, setLogs] = React.useState([]);
      const [wordOpen, setWordOpen] = React.useState(false);
      const [mappingView, setMappingView] = React.useState(null);
      const [agentSessionId, setAgentSessionId] = React.useState(() => {
        try {
          return window.localStorage.getItem('dsh-word-agent-session') || null;
        } catch {
          return null;
        }
      });
      const [zoom, setZoom] = React.useState(null);
      const logRef = React.useRef(null);
      const chatRef = React.useRef(null);

      React.useEffect(() => {
        if (chatRef.current) chatRef.current.scrollTop = chatRef.current.scrollHeight;
      }, [messages, busy]);

      // Host apps sometimes intercept file drops on window/document. Capture the
      // drop here so a file dropped on an upload card always reaches that card.
      React.useEffect(() => {
        const zoneOf = (target) =>
          target && target.closest ? target.closest('.wp-drop[data-drop-key]') : null;
        const onDragOver = (event) => {
          if (!zoneOf(event.target)) return;
          event.preventDefault();
          if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
        };
        const onDrop = (event) => {
          const zone = zoneOf(event.target);
          if (!zone) return;
          event.preventDefault();
          event.stopPropagation();
          const handler = dropTargets.get(zone.getAttribute('data-drop-key'));
          const files = event.dataTransfer && event.dataTransfer.files;
          if (handler && files && files[0]) handler(files[0]);
        };
        window.addEventListener('dragover', onDragOver, true);
        window.addEventListener('drop', onDrop, true);
        return () => {
          window.removeEventListener('dragover', onDragOver, true);
          window.removeEventListener('drop', onDrop, true);
        };
      }, []);

      React.useEffect(() => {
        if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
      }, [logs]);

      React.useEffect(() => {
        if (!zoom) return undefined;
        const onKey = (event) => {
          if (event.key === 'Escape') setZoom(null);
        };
        window.addEventListener('keydown', onKey);
        return () => window.removeEventListener('keydown', onKey);
      }, [zoom]);

      React.useEffect(() => {
        try {
          const raw = window.localStorage.getItem('dsh-word-files-v1');
          if (!raw) return;
          const saved = JSON.parse(raw);
          if (saved && (saved.source || saved.template)) {
            setFiles({ source: saved.source || null, template: saved.template || null });
          }
        } catch {
          /* ignore a broken cache */
        }
      }, []);

      React.useEffect(() => {
        try {
          window.localStorage.setItem(
            'dsh-word-files-v1',
            JSON.stringify({ source: files.source, template: files.template }),
          );
        } catch {
          /* storage may be unavailable */
        }
      }, [files]);

      async function send(text) {
        const value = String(text ?? input).trim();
        if (!value || busy) return;
        const delegated = await promptAgent(
          `${taskBrief('作为 Word 流水线 Agent 继续处理：')}\n用户要求：${value}`,
        );
        if (delegated) {
          setInput('');
          setMessages((current) => [
            ...current,
            { role: 'user', text: value },
            { role: 'agent', text: '已发给 Agent，请在新对话里继续。' },
          ]);
          return;
        }
        setInput('');
        setMessages((current) => [...current, { role: 'user', text: value }]);
        setBusy(true);
        setError('');
        try {
          const payload = {
            text: value,
            history: messages
              .filter((message) => message.role === 'user')
              .slice(-6)
              .map((message) => ({ role: 'user', text: message.text })),
            template: files.template && files.template.path,
            source: files.source && files.source.path,
          };
          let data = await call('/agent/chat', payload);
          let note = '';
          if (data.fallback) {
            const local = await call('/agent/instruct', {
              text: value,
              template: payload.template,
              source: payload.source,
            });
            data = { action: local.action || 'none', reply: local.reply, result: local.result };
            note = '\n（模型未接入，已用本地指令解析）';
          }
          setMessages((current) => [
            ...current,
            { role: 'agent', text: `${data.reply || '(无回复)'}${note}` },
          ]);
          if (Array.isArray(data.logs)) {
            setLogs((current) => [...current.slice(-150), ...data.logs]);
          }
          if (data.action === 'convert' && data.result) {
            setOutcome(data.result);
            if (Array.isArray(data.result.steps)) setSteps(data.result.steps);
            if (data.result.ok) {
              setPreview({ path: data.result.png || data.result.docx, version: Date.now() });
            } else if (data.result.error) {
              setError(data.result.error);
            }
          } else if (data.result) {
            if (data.result.steps && data.result.steps.length) setSteps(data.result.steps);
            const artifacts = data.result.artifacts || [];
            const png = [...artifacts].reverse().find((item) => item.endsWith('.png'));
            const docx = [...artifacts].reverse().find((item) => item.endsWith('.docx'));
            if (png) setPreview({ path: png, version: Date.now() });
            else if (docx) setPreview({ path: docx, version: Date.now() });
          }
        } catch (failure) {
          setMessages((current) => [
            ...current,
            { role: 'agent', text: `出错：${failure.message || failure}` },
          ]);
        } finally {
          setBusy(false);
        }
      }

      async function pullProgress() {
        try {
          const response = await fetch(API + '/progress', { headers: { 'X-Word-Token': token() } });
          const data = await response.json();
          if (data && data.ok && Array.isArray(data.steps)) setLiveCount(data.steps.length);
          if (data && data.preview) {
            setPreview((current) =>
              current && current.path === data.preview
                ? current
                : { path: data.preview, version: Date.now() },
            );
          }
        } catch {
          /* the next tick retries */
        }
      }

      async function openWord() {
        try {
          const target = (outcome && outcome.docx) || 'fixtures/build/out.docx';
          const data = await call('/open-word', { path: target });
          setWordOpen(Boolean(data.ok));
          if (!data.ok) setError(data.error || '无法打开 Word');
          return data;
        } catch (failure) {
          setError(String(failure.message || failure));
          return null;
        }
      }

      async function showMapping() {
        try {
          const response = await fetch(API + '/mapping', { headers: { 'X-Word-Token': token() } });
          const data = await response.json();
          setMappingView(data.debug || { final: [], attempts: [] });
        } catch (failure) {
          setError(String(failure.message || failure));
        }
      }

      function delegateToAgent() {
        const brief = [
          '请作为 Word 流水线 Agent 处理这次转换：',
          `- 模板：${(files.template && files.template.path) || '未上传（使用示例模板）'}`,
          `- 资料：${(files.source && files.source.path) || '未上传（使用示例资料）'}`,
          outcome && outcome.checkIssues && outcome.checkIssues.length
            ? `- 上次检查问题：${outcome.checkIssues.map((item) => `[${item.rule}] ${item.message}`).join('；')}`
            : '- 上次检查：无记录',
          '要求：先读 docs/word-pipeline-skill.md（命令契约），再读模板理解章节（标题/作者/摘要/正文/参考文献），按章节从资料取内容，原地替换并保留模板格式；',
          '缺内容时记录并继续；完成后运行 python -m dsh_word check，并修复格式差异。',
        ].join('\n');
        try {
          if (navigator.clipboard) navigator.clipboard.writeText(brief).catch(() => {});
        } catch {
          /* clipboard is optional */
        }
        if (uiWorkspace && typeof uiWorkspace.startSession === 'function') {
          uiWorkspace.startSession(undefined, { prompt: brief, whenDraftExists: 'replace' });
          setMessages((current) => [
            ...current,
            { role: 'agent', text: '已把任务交给 Agent：新对话里预填了模板/资料路径与要求，发送即可让 Agent 直接处理文件。' },
          ]);
        } else {
          setMessages((current) => [
            ...current,
            { role: 'agent', text: '任务说明已复制到剪贴板，请粘贴到对话里发送。' },
          ]);
        }
      }

      function taskBrief(title) {
        const template = files.template && files.template.path ? files.template.path : '（未上传，使用示例模板）';
        const source = files.source && files.source.path ? files.source.path : '（未上传，使用示例资料）';
        const issues = outcome && outcome.checkIssues && outcome.checkIssues.length
          ? outcome.checkIssues.map((item) => `[${item.rule}] ${item.message}`).join('；')
          : '无';
        return [
          title,
          `- 模板：${template}`,
          `- 资料：${source}`,
          `- 上次检查问题：${issues}`,
          '要求：先读 docs/word-pipeline-skill.md（命令契约），再读模板理解章节（标题/作者/摘要/正文/参考文献），按章节从资料取内容，原地替换并保留模板格式；',
          '缺内容时记录并继续；完成后运行 python -m dsh_word check，并修复格式差异。',
        ].join('\n');
      }

      async function promptAgent(text) {
        if (!clientSessions || typeof clientSessions.create !== 'function') return false;
        try {
          let id = agentSessionId;
          if (!id) {
            try {
              id = await clientSessions.create({ agentPreset: 'word-master' });
            } catch {
              id = await clientSessions.create({});
            }
            setAgentSessionId(id);
            try {
              window.localStorage.setItem('dsh-word-agent-session', id);
            } catch {
              /* storage is optional */
            }
          }
          if (uiWorkspace && typeof uiWorkspace.openSession === 'function') uiWorkspace.openSession(id);
          await clientSessions.using(id, { source: 'controllerOperation' }, async (reference) => {
            await reference.ready;
            await reference.binding.session.prompt([{ type: 'text', text }], 'queue');
          });
          return true;
        } catch (failure) {
          setError(String(failure.message || failure));
          return false;
        }
      }

      async function convert() {
        if (busy) return;
        const delegated = await promptAgent(
          taskBrief('请作为 Word 流水线 Agent 处理这次转换（使用 Word Master 工具链）：'),
        );
        if (delegated) {
          setMessages((current) => [
            ...current,
            { role: 'agent', text: '已创建 Word Master 会话：模板、资料和任务说明已发过去，Agent 会直接开始处理。' },
          ]);
          return;
        }
        setBusy(true);
        setError('');
        setOutcome(null);
        setSteps([]);
        setLogs([]);
        setPreview(null);
        setLiveCount(0);
        setStageText('读取资料…');
        try {
          const data = await streamConvert(
            {
              template: files.template && files.template.path,
              source: files.source && files.source.path,
            },
            (event) => {
              if (event.type === 'stage') {
                setStageText(
                  event.name === 'fill'
                    ? '读取资料并填充…'
                    : event.name === 'check'
                      ? '检查中…'
                      : event.name === 'render'
                        ? '渲染中…'
                        : '完成',
                );
              } else if (event.type === 'step') {
                setSteps((current) => [...current, event.item]);
                setStageText(`填充 ${event.item.op}：${event.item.slot || event.item.kind}`);
              } else if (event.type === 'preview') {
                setPreview({ path: event.path, version: Date.now() });
              } else if (event.type === 'log') {
                setLogs((current) => [...current.slice(-199), event.entry]);
              } else if (event.type === 'missing') {
                setMessages((current) => [
                  ...current,
                  {
                    role: 'agent',
                    text: `资料里没有找到「${event.segment}」可用的内容，先跳过这部分，继续处理后面的章节。`,
                  },
                ]);
              }
            },
          );
          setOutcome(data);
          if (data.steps && data.steps.length) setSteps(data.steps);
          if (data.ok) setPreview({ path: data.png || data.docx, version: Date.now() });
        } catch (failure) {
          setError(String(failure.message || failure));
        } finally {
          setStageText('');
          setBusy(false);
        }
      }

      const downloadHref = (value) =>
        `${API}/download?path=${encodeURIComponent(value)}&t=${encodeURIComponent(token())}`;

      const stageRow = (key, label, state, meta) =>
        h(
          'div',
          { key, className: 'wp-stage-row', 'data-state': state },
          h('span', { className: 'wp-stage-dot' }),
          h('span', null, label),
          meta ? h('span', { className: 'wp-stage-meta' }, meta) : null,
        );

      const translateStage = { fill: '填充', check: '检查', render: '渲染' };

      const outcomeBlock = busy
        ? h(
            'div',
            { className: 'wp-progress' },
            h('div', { className: 'wp-subhead' }, '正在转换'),
            stageRow('fill', '填充', 'run', liveCount ? `${liveCount} 个中间版本` : ''),
            stageRow('check', '检查', 'idle', ''),
            stageRow('render', '渲染', 'idle', ''),
          )
        : outcome
          ? h(
              'div',
              { className: 'wp-progress' },
              h('div', { className: 'wp-subhead' }, outcome.ok ? '转换完成' : '转换中断'),
              ...(outcome.stages || []).map((stage) =>
                stageRow(
                  stage.name,
                  translateStage[stage.name] || stage.name,
                  stage.ok ? 'ok' : 'fail',
                  stage.steps ? `${stage.steps} 步` : '',
                ),
              ),
              !outcome.ok && outcome.error ? h('div', { className: 'wp-error' }, outcome.error) : null,
              outcome.ok
                ? h(
                    'div',
                    { className: 'wp-dl-row' },
                    h(
                      'a',
                      {
                        className: 'wp-dl wp-dl-primary',
                        href: downloadHref(outcome.docx),
                        download: '',
                      },
                      '下载 DOCX',
                    ),
                    h(
                      'a',
                      { className: 'wp-dl', href: downloadHref(outcome.pdf), download: '' },
                      '下载 PDF',
                    ),
                  )
                : null,
              outcome.checkIssues && outcome.checkIssues.length
                ? h(
                    'div',
                    { className: 'wp-check-issues' },
                    ...outcome.checkIssues.slice(0, 8).map((issue, index) =>
                      h(
                        'div',
                        { key: index, className: 'wp-check-issue' },
                        `[${issue.severity || 'error'}] [${issue.rule}] ${issue.message}`,
                      ),
                    ),
                    outcome.checkIssues.length > 8
                      ? h(
                          'div',
                          { className: 'wp-check-issue' },
                          `…共 ${outcome.checkIssues.length} 项`,
                        )
                      : null,
                  )
                : null,
            )
          : null;

      const workspace = h(
        'div',
        { className: 'wp-workspace' },
        h(
          'section',
          { className: 'wp-panel' },
          h(
            'header',
            { className: 'wp-panel-head' },
            h('span', { className: 'wp-panel-title' }, '一键转换'),
            h(
              'span',
              { className: 'wp-panel-meta' },
              busy ? '执行中…' : '填充 → 检查 → 渲染',
            ),
          ),
          h(
            'div',
            { className: 'wp-panel-body' },
            h(
              'div',
              { className: 'wp-grid' },
              h(DropZone, {
                label: '资料 DOCX',
                hint: '拖入资料，或留空用示例',
                accept: '.docx',
                kind: 'source',
                value: files.source,
                onChange: (value) => setFiles((current) => ({ ...current, source: value })),
                onError: setError,
              }),
              h(DropZone, {
                label: '模板 DOCX',
                hint: '拖入模板，或留空用示例',
                accept: '.docx',
                kind: 'template',
                value: files.template,
                onChange: (value) => setFiles((current) => ({ ...current, template: value })),
                onError: setError,
              }),
            ),
            h(
              'button',
              { className: 'wp-cta', disabled: busy, onClick: convert },
              busy ? '正在转换…' : '开始转换',
            ),
            h(
              'div',
              { className: 'wp-hint' },
              busy && stageText ? stageText : '上传资料与模板后点一次，自动完成填充、检查与预览。',
            ),
            outcomeBlock,
            h('div', { className: 'wp-divider' }),
            h('div', { className: 'wp-subhead' }, '需要调整？直接说明，例如「把图改成 10cm」'),
            h(
              'div',
              { className: 'wp-chat wp-chat-compact', ref: chatRef },
              ...messages.map((message, index) =>
                h('div', { key: index, className: 'wp-msg', 'data-role': message.role }, message.text),
              ),
              busy
                ? null
                : null,
            ),
            h(
              'div',
              { className: 'wp-input-row' },
              h('input', {
                className: 'wp-input',
                value: input,
                disabled: busy,
                placeholder: '例如：把图改成 10cm',
                onChange: (event) => setInput(event.target.value),
                onKeyDown: (event) => {
                  if (event.key === 'Enter') send();
                },
              }),
              h(
                'button',
                { className: 'wp-send', disabled: busy, onClick: () => send() },
                busy ? '…' : '调整',
              ),
            ),
          ),
        ),
        h(
          'section',
          { className: 'wp-panel' },
          h(
            'header',
            { className: 'wp-panel-head' },
            h('span', { className: 'wp-panel-title' }, '文档预览'),
            h(
              'span',
              { className: 'wp-panel-meta' },
              busy ? stageText || '运行中' : preview ? 'Word 渲染' : '等待运行',
            ),
          ),
          h(
            'div',
            { className: 'wp-panel-body wp-preview-body' },
            h(
              'div',
              { className: 'wp-word' },
              h('span', { className: 'wp-word-icon' }, h(AgentIcon, { size: 30 })),
              h(
                'div',
                { className: 'wp-word-copy' },
                h('div', { className: 'wp-word-title' }, wordOpen ? 'Word 正在实时显示' : '在 Word 中查看并编辑'),
                h(
                  'div',
                  { className: 'wp-word-sub' },
                  wordOpen
                    ? '每一步都会写入 Word，并高亮刚修改的段落。'
                    : '打开 Word 后，转换过程会实时同步到文档里，改动处用黄色标出。',
                ),
              ),
              h(
                'button',
                { className: 'wp-word-open', onClick: openWord },
                wordOpen ? '重新打开 Word' : '在 Word 中打开',
              ),
              h('button', { className: 'wp-word-map', onClick: showMapping }, '查看映射详情'),
              h('button', { className: 'wp-word-map', onClick: delegateToAgent }, '交给 Agent 处理'),
            ),
            steps.length
              ? h(
                  'div',
                  { className: 'wp-steplist' },
                  ...steps.map((item) =>
                    h(
                      'button',
                      {
                        key: `${item.op}:${item.step}`,
                        className: 'wp-step',
                        'data-active': preview && preview.path === item.step ? '1' : '0',
                        onClick: () => setPreview({ path: item.step, version: Date.now() }),
                      },
                      h('span', { className: 'wp-step-num' }, String(item.op).padStart(2, '0')),
                      `${item.kind} · ${item.slot}`,
                    ),
                  ),
                )
              : null,
            error ? h('div', { className: 'wp-error' }, error) : null,
            h(
              'div',
              { className: 'wp-log' },
              h(
                'div',
                { className: 'wp-log-head' },
                h('span', null, 'Agent 日志'),
                h('span', { className: 'wp-log-count' }, `${logs.length} 条`),
              ),
              h(
                'div',
                { className: 'wp-log-body', ref: logRef },
                logs.length
                  ? logs.map((entry, index) =>
                      h(
                        'div',
                        { key: index, className: 'wp-log-row', 'data-level': entry.level || 'info' },
                        h(
                          'span',
                          { className: 'wp-log-time' },
                          new Date(entry.t).toLocaleTimeString('zh-CN', { hour12: false }),
                        ),
                        h('span', { className: 'wp-log-phase' }, entry.phase || ''),
                        h('span', { className: 'wp-log-msg' }, entry.message || ''),
                        entry.ms ? h('span', { className: 'wp-log-ms' }, `${entry.ms}ms`) : null,
                      ),
                    )
                  : h('div', { className: 'wp-log-empty' }, '还没有运行记录，点一次「开始转换」就能看到每一步。'),
              ),
            ),
          ),
        ),
      );

      const gallery = h(
        'div',
        { className: 'wp-list' },
        h(
          'button',
          { className: 'wp-row', onClick: () => setView('workspace') },
          h('span', { className: 'wp-row-icon' }, h(AgentIcon, { size: 30 })),
          h(
            'div',
            { className: 'wp-row-text' },
            h('div', { className: 'wp-row-title' }, 'Word 流水线助手'),
            h(
              'div',
              { className: 'wp-row-desc' },
              '一键完成填充、检查与 Word 渲染预览，成品可直接下载；需要时再用对话微调。',
            ),
          ),
          h('span', { className: 'wp-row-meta' }, '本地 · python'),
          h('span', { className: 'wp-row-arrow' }, '→'),
        ),
        h(
          'div',
          { className: 'wp-row-note' },
          '每个 Agent 都是一个独立应用：一个卡片、一套自己的端点与界面。更多 Agent 按同样方式接入。',
        ),
      );

      return h(
        'div',
        { className: 'wp-page' },
        h('style', null, CSS),
        h(
          'div',
          { className: 'wp-page-inner' },
          h(
            'div',
            { className: 'wp-hero-band' },
            h(
              'div',
              { className: 'wp-hero' },
              h(
                'div',
                { className: 'wp-hero-left' },
                h(
                  'span',
                  { className: 'wp-logo' },
                  view === 'workspace' ? h(AgentIcon, { size: 40 }) : h(ManagerIcon, { size: 40 }),
                ),
                h(
                  'div',
                  { style: { minWidth: 0 } },
                  h(
                    'div',
                    { className: 'wp-hero-title' },
                    view === 'workspace' ? 'Word 流水线助手' : 'Agent 管理器',
                  ),
                  h(
                    'div',
                    { className: 'wp-hero-sub' },
                    view === 'workspace'
                      ? '资料 + 模板 → 一键转换；需要时再调整'
                      : '每个 Agent 是一个独立应用，有自己的界面和执行后端',
                  ),
                ),
              ),
              h(
                'div',
                { style: { display: 'flex', alignItems: 'center', gap: '8px' } },
                h(
                  'span',
                  { className: 'wp-pill' },
                  h('span', { className: 'wp-pill-dot', 'data-busy': busy ? '1' : '0' }),
                  busy ? '执行中' : '就绪',
                ),
                view === 'workspace'
                  ? h('button', { className: 'wp-btn', onClick: () => setView('grid') }, '← 全部 Agent')
                  : null,
              ),
            ),
          ),
          view === 'workspace' ? workspace : gallery,
        ),
        zoom
          ? h(
              'div',
              { className: 'wp-lightbox', onClick: () => setZoom(null) },
              h(
                'button',
                {
                  className: 'wp-lightbox-close',
                  title: '关闭（Esc）',
                  onClick: (event) => {
                    event.stopPropagation();
                    setZoom(null);
                  },
                },
                '×',
              ),
              h('img', {
                alt: '文档预览',
                src: `${API}/preview?path=${encodeURIComponent(zoom.path)}&t=${encodeURIComponent(token())}&v=${zoom.version}`,
                onClick: (event) => event.stopPropagation(),
              }),
            )
          : null,
        mappingView
          ? h(
              'div',
              { className: 'wp-lightbox', onClick: () => setMappingView(null) },
              h(
                'div',
                { className: 'wp-map-modal', onClick: (event) => event.stopPropagation() },
                h(
                  'div',
                  { className: 'wp-map-head' },
                  h('span', null, '结构映射详情'),
                  h(
                    'button',
                    {
                      className: 'wp-lightbox-close',
                      onClick: () => setMappingView(null),
                    },
                    '×',
                  ),
                ),
                h(
                  'div',
                  { className: 'wp-map-body' },
                  mappingView.notes
                    ? h('div', { className: 'wp-map-notes' }, mappingView.notes)
                    : null,
                  (mappingView.final || []).length
                    ? h(
                        'div',
                        { className: 'wp-map-list' },
                        ...mappingView.final.map((item, index) =>
                          h(
                            'div',
                            { key: index, className: 'wp-map-row' },
                            h('span', { className: 'wp-map-t' }, item.template || ''),
                            h('span', { className: 'wp-map-arrow' }, '←'),
                            h('span', { className: 'wp-map-s' }, item.source || '—'),
                            h('span', { className: 'wp-map-a' }, item.action || 'replace'),
                          ),
                        ),
                      )
                    : h('div', { className: 'wp-map-notes' }, '还没有映射记录，先转换一次。'),
                  (mappingView.verify &&
                    mappingView.verify.issues &&
                    mappingView.verify.issues.length)
                    ? h(
                        'div',
                        { className: 'wp-map-issues' },
                        ...mappingView.verify.issues.map((issue, index) =>
                          h(
                            'div',
                            { key: index, className: 'wp-map-issue', 'data-level': issue.level },
                            `${issue.level}：${issue.message}`,
                          ),
                        ),
                      )
                    : null,
                  ...(mappingView.attempts || []).map((attempt, index) =>
                    h(
                      'details',
                      { key: `raw-${index}`, className: 'wp-map-raw' },
                      h('summary', null, `模型原始输出（${attempt.stage}）`),
                      h('pre', null, attempt.raw || ''),
                    ),
                  ),
                ),
              ),
            )
          : null,
      );
    }

    let tabsReady = false;

    return {
      inject: ['slots'],
      apply(ctx) {
        runtime = ctx;
        try {
          if (typeof ctx.inject === 'function') {
            ctx.inject(['sessions'], (scope) => {
              clientSessions = scope.sessions || null;
            });
          }
          clientSessions = clientSessions || ctx.get('sessions') || null;
        } catch {
          clientSessions = null;
        }
        try {
          if (typeof ctx.inject === 'function') {
            ctx.inject(['uiWorkspace'], (scope) => {
              uiWorkspace = scope.uiWorkspace || null;
            });
          }
          uiWorkspace = uiWorkspace || ctx.get('uiWorkspace') || null;
        } catch {
          uiWorkspace = null;
        }
        const tabs = ctx.get('sidebarRightTabs');
        if (tabs && typeof tabs.register === 'function') {
          ctx.effect(() =>
            tabs.register({
              id: 'word-pipeline',
              kind: 'word-pipeline',
              title: () => 'Word 流水线',
              guide: [
                {
                  id: 'word-pipeline',
                  title: 'Word 流水线',
                  description: '资料填充、检查、合批、渲染',
                },
              ],
            }),
          );
          tabsReady = true;
        }
        ctx.slots.inject('sidebar.right.pane.tab', () =>
          ctx.slots.register({ name: 'sidebar.right.pane.tab', key: 'word-pipeline' }, Panel),
        );
        ctx.slots.inject('shell.overlay', () =>
          ctx.slots.register(
            { name: 'shell.overlay', id: 'word-pipeline-overlay', order: 50 },
            OverlayHost,
          ),
        );
        ctx.slots.inject('sidebar.panellist', () =>
          ctx.slots.register(
            { name: 'sidebar.panellist', id: 'agents', order: 2, label: () => 'Agent 管理器' },
            ManagerIcon,
          ),
        );
        setTimeout(async () => {
          try {
            const response = await fetch(API + '/state', { headers: { 'X-Word-Token': token() } });
            const state = await response.json();
            if (!state || !state.autoPanel) return;
            const layout = ctx.get('layout');
            if (layout && typeof layout.selectPanel === 'function') layout.selectPanel('agents');
          } catch {
            /* startup nicety only */
          }
        }, 1500);

        ctx.slots.inject('main', () =>
          ctx.slots.register({ name: 'main', key: 'agents' }, AgentPanelPage),
        );
      },
    };
  },
});
