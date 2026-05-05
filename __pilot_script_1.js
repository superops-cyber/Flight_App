
(function(){
  var prev = document.getElementById('tab2-briefing-styles');
  if (prev) prev.remove();
  var s = document.createElement('style');
  s.id = 'tab2-briefing-styles';
  s.textContent = `
#briefing-content {
  background: linear-gradient(180deg, rgba(26,42,36,0.88) 0%, rgba(26,42,36,0.18) 220px),
              repeating-linear-gradient(0deg, rgba(0,0,0,0.07) 0 1px, transparent 1px 110px),
              repeating-linear-gradient(90deg, rgba(0,0,0,0.06) 0 1px, transparent 1px 110px),
              linear-gradient(180deg, #4b5645 0%, #7b7d66 40%, #8f927c 100%);
  font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
  min-height: 100vh;
  padding: 0 8px 105px 8px;
  width: 100%;
  max-width: 100%;
  overflow-x: hidden;
  box-sizing: border-box;
}
#briefing-content * { font-weight: 800; box-sizing: border-box; }
#briefing-content .brief-main { min-width: 0; padding-top: 8px; width: 100%; max-width: 100%; overflow-x: hidden; }
#briefing-content .brief-hero {
  margin: 0 -8px 8px -8px;
  padding: 10px 12px 8px 12px;
  background: linear-gradient(180deg, #2d69bd 0%, #1f4d8d 100%);
  border-bottom: 2px solid #143764;
  box-shadow: 0 4px 10px rgba(0,0,0,0.22);
}
#briefing-content .brief-title { font-size: 2rem; color: #fff; text-transform: uppercase; margin: 0; line-height: 1.05; }
#briefing-content .brief-subtitle { display: none; }
#briefing-content .brief-info {
  margin-top: 8px;
  background: linear-gradient(180deg, rgba(232,223,201,0.97), rgba(212,201,177,0.95));
  border: 1px solid #8d7e64;
  border-radius: 8px;
  padding: 8px 10px;
  color: #111;
  font-size: 1.17rem;
}
#briefing-content .stats-row {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 4px;
  margin-top: 8px;
  width: 100%;
  max-width: 100%;
  overflow-x: hidden;
}
#briefing-content .stat-card {
  background: linear-gradient(180deg, #3a3a3a 0%, #2a2a2a 100%) !important;
  border: 1px solid #6e6e6e !important;
  border-top: 1px solid #6e6e6e !important;
  border-radius: 8px;
  padding: 6px 4px;
  text-align: center;
  box-shadow: 0 3px 8px rgba(0,0,0,0.2);
  min-width: 0;
}
#briefing-content .stat-val { display: block; font-size: 1.42rem; color: #8ce26c !important; line-height: 1.1; background: transparent !important; }
#briefing-content .stat-val.green,
#briefing-content .stat-val.blue { color: #8ce26c !important; background: transparent !important; }
#briefing-content .stat-label { display: block; font-size: 0.78rem; color: #c7c7c7; text-transform: uppercase; margin-top: 3px; }
#briefing-content .brief-section-title { margin: 6px 0 6px 0; font-size: 1.38rem; text-transform: uppercase; color: #f0a03b; background: transparent !important; }
#briefing-content .brief-section-title.green { color: #8dcf8f; background: transparent !important; }
#briefing-content .brief-panel { background: transparent; border: none; box-shadow: none; padding: 0; margin-top: 6px; width: 100%; max-width: 100%; overflow-x: hidden; }

#briefing-content .leg-card {
  margin-bottom: 8px;
  background: linear-gradient(180deg, #ffffff 0%, #ececec 100%);
  border: 1px solid #a7a7a7;
  border-radius: 10px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.2);
  overflow: hidden;
}
#briefing-content .leg-card.is-complete { opacity: 0.72; }
#briefing-content .leg-head {
  background: #fff;
  padding: 8px 10px 6px 10px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 8px;
}
#briefing-content .leg-route-line { font-size: 1.35rem; color: #111; text-transform: uppercase; line-height: 1.15; }
#briefing-content .leg-via-line { font-size: 0.99rem; color: #333; text-transform: uppercase; margin-top: 2px; }
#briefing-content .leg-select-btn { background: #101010; color: #fff; border: none; border-radius: 7px; padding: 4px 8px; font-size: 0.93rem; text-transform: uppercase; white-space: nowrap; }
#briefing-content .leg-select-btn.complete { background: #757575; }
#briefing-content .leg-body { padding: 6px 10px 8px 10px; display: flex; flex-wrap: wrap; gap: 8px; align-items: flex-start; }
#briefing-content .leg-summary { flex: 1 1 320px; min-width: 0; }
#briefing-content .brief-chip-row { display: flex; gap: 4px; flex-wrap: wrap; margin-bottom: 4px; }
#briefing-content .brief-badge { padding: 3px 6px; border-radius: 6px; color: #fff !important; font-size: 0.87rem; text-transform: uppercase; display: inline-block; line-height: 1; }
#briefing-content .brief-badge.red { background: #D32F2F !important; }
#briefing-content .brief-badge.green { background: #388E3C !important; }
#briefing-content .brief-badge.blue { background: #2167D1 !important; }
#briefing-content .brief-badge.gray { background: #111 !important; }
#briefing-content .brief-badge.cache-draw-alert {
  background: linear-gradient(180deg, #ffd54f 0%, #f57f17 100%) !important;
  color: #1a1a1a !important;
  border: 1px solid #8b5e00;
  box-shadow: 0 0 0 1px rgba(255, 193, 7, 0.35), 0 0 10px rgba(255, 152, 0, 0.55);
  animation: cachePulse 1.2s ease-in-out infinite;
}
@keyframes cachePulse {
  0% { transform: scale(1); }
  50% { transform: scale(1.04); }
  100% { transform: scale(1); }
}
#briefing-content .brief-data-line { font-size: 0.96rem; color: #111; line-height: 1.2; text-transform: none; }
#briefing-content .brief-data-line .weight-limit-note {
  color: #7a1a1a;
  font-weight: 900;
}
#briefing-content .brief-chip-panel.cache-usage-alert {
  margin-top: 8px;
  border: 2px solid #b71c1c;
  border-radius: 8px;
  padding: 6px 8px;
  background:
    repeating-linear-gradient(-45deg, rgba(255, 235, 238, 0.75) 0 8px, rgba(255, 205, 210, 0.75) 8px 16px),
    linear-gradient(180deg, #ffebee 0%, #ffcdd2 100%);
  box-shadow: 0 0 0 1px rgba(183, 28, 28, 0.35), 0 0 12px rgba(183, 28, 28, 0.45);
  animation: cacheUsagePulse 0.9s ease-in-out infinite;
}
#briefing-content .cache-usage-title {
  font-size: 0.72rem;
  color: #7f0000;
  letter-spacing: 0.05em;
  margin-bottom: 3px;
}
#briefing-content .cache-usage-row {
  font-size: 0.93rem;
  color: #2a0d0d;
  margin-bottom: 5px;
}
#briefing-content .cache-usage-row .qty {
  font-size: 1.06rem;
  font-weight: 900;
  color: #b71c1c;
}
#briefing-content .cache-status-pill {
  display: inline-block;
  padding: 4px 8px;
  border-radius: 999px;
  font-size: 0.75rem;
  letter-spacing: 0.03em;
  border: 1px solid transparent;
}
#briefing-content .cache-status-pill.pending {
  background: #c62828;
  color: #fff;
  border-color: #8e0000;
}
#briefing-content .cache-status-pill.verified {
  background: #2e7d32;
  color: #fff;
  border-color: #1b5e20;
}
@keyframes cacheUsagePulse {
  0% { box-shadow: 0 0 0 1px rgba(183, 28, 28, 0.35), 0 0 8px rgba(183, 28, 28, 0.35); }
  50% { box-shadow: 0 0 0 1px rgba(183, 28, 28, 0.55), 0 0 16px rgba(183, 28, 28, 0.65); }
  100% { box-shadow: 0 0 0 1px rgba(183, 28, 28, 0.35), 0 0 8px rgba(183, 28, 28, 0.35); }
}
#briefing-content .brief-data-line .good { color: #388E3C; background: transparent !important; }
#briefing-content .pax-block { background: rgba(255,255,255,0.52); border: 1px solid #c7c7c7; border-radius: 8px; padding: 6px; min-width: 0; width: 100%; flex: 1 1 100%; overflow-x: auto; }
#briefing-content .pax-row {
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: nowrap;
  white-space: nowrap;
  border-bottom: 1px solid rgba(0,0,0,0.08);
  padding: 2px 0;
}
#briefing-content .pax-row:last-child { border-bottom: none; }
#briefing-content .pax-col { color: #222; font-size: 0.91rem; }
#briefing-content .pax-col.name { min-width: 120px; }
#briefing-content .pax-col.sex { min-width: 34px; text-align: center; }
#briefing-content .pax-col.cat { min-width: 58px; text-align: center; }
#briefing-content .pax-col.weight { min-width: 90px; text-align: right; }
#briefing-content .pax-col.cargo { min-width: 95px; text-align: right; }
#briefing-content .pax-name { color: #111; font-size: 1rem; }
#briefing-content .pax-name .good { color: #388E3C; background: transparent !important; }
#briefing-content .pax-meta { display: flex; gap: 3px; flex-wrap: wrap; margin-top: 3px; }
#briefing-content .pax-tag { font-size: 0.78rem; padding: 2px 5px; border-radius: 10px; text-transform: uppercase; color: #fff; }
#briefing-content .tag-m, #briefing-content .tag-f { background: #c53333; }
#briefing-content .tag-age { background: #6e6e6e; }
#briefing-content .pax-load { font-size: 0.87rem; color: #444; margin-top: 3px; text-transform: none; }

#briefing-content .leg-inputs {
  margin-top: 0;
  background: transparent;
  border: none;
  border-radius: 0;
  display: flex;
  flex-wrap: nowrap;
  gap: 6px;
  overflow: visible;
  align-items: flex-start;
  width: auto;
  flex: 0 0 auto;
}
#briefing-content .fieldbox { padding: 0; border-right: none; min-width: 0; flex: 0 0 auto; }
#briefing-content .fieldbox label,
#briefing-content .brief-entry-box label,
#briefing-content .brief-tank label,
#briefing-content .brief-total-box label {
  display: block;
  font-size: 0.78rem;
  color: #666;
  text-transform: uppercase;
  margin-bottom: 3px;
}
#briefing-content .fieldbox input,
#briefing-content .brief-entry-box input,
#briefing-content .brief-tank input {
  width: 100%;
  height: 28px;
  box-sizing: border-box;
  border: none;
  background: transparent;
  color: #111;
  font-size: 1.32rem;
  padding: 0;
  border-radius: 0;
}
#briefing-content .leg-inputs .fieldbox label { margin-bottom: 2px; font-size: 0.68rem; color: #6a6a6a; line-height: 1; }
#briefing-content .leg-inputs .fieldbox input.plan-id-input,
#briefing-content .leg-inputs .fieldbox input.zulu-input {
  font-size: 14px;
  font-family: "SFMono-Regular", Menlo, Consolas, monospace;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  border: 1px solid #a9b4bf;
  border-radius: 5px;
  background: #fff;
  padding: 0 6px;
  height: 26px;
}
#briefing-content .leg-inputs .fieldbox input.plan-id-input {
  width: 12.8ch;
}
#briefing-content .leg-inputs .fieldbox input.zulu-input {
  width: 6.8ch;
  text-align: center;
}
#briefing-content .fieldbox.no-plan-box { display: flex; flex-direction: column; align-items: flex-start; }
#briefing-content .fieldbox.no-plan-box label { display: block; }
#briefing-content .no-plan-wrap { display: flex; align-items: center; justify-content: center; gap: 0; height: 26px; min-width: 18px; }
#briefing-content .no-plan-wrap input[type="checkbox"] {
  width: 16px;
  height: 16px;
  margin: 0;
  border: revert;
  background: revert;
  appearance: checkbox;
  -webkit-appearance: checkbox;
  position: static;
  opacity: 1;
  pointer-events: auto;
  visibility: visible;
  display: block;
  accent-color: #1f6cd1;
}
#briefing-content .no-plan-wrap span { display: none; }
#briefing-content .brief-entry-box input.volts-example { color: #95a0ad; }
#briefing-content .oil-choice-row { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 4px; height: 32px; min-width: 0; }
#briefing-content .oil-choice {
  border: 1px solid #6a6a6a;
  background: #202020;
  color: #d4d4d4;
  border-radius: 6px;
  font-size: 0.98rem;
  font-weight: 800;
  min-width: 0;
  width: 100%;
  padding: 0;
  line-height: 1;
  overflow: hidden;
}
#briefing-content .oil-choice.active {
  background: #1f6cd1;
  border-color: #3176d4;
  color: #fff;
}
#briefing-content .brief-tank-btn {
  width: 100%;
  height: 28px;
  border: 1px solid #4f6380;
  background: #1d2936;
  color: #fff;
  border-radius: 6px;
  font-size: 1.2rem;
  font-weight: 800;
}
#briefing-content .brief-tank-btn.startup-selected {
  background: #1a3a2a;
  border-color: #43a047;
  color: #dff7cf;
  box-shadow: inset 0 0 0 1px rgba(140, 226, 108, 0.35);
}

#briefing-content .brief-entry-grid {
  display: grid;
  grid-template-columns: repeat(8, 1fr);
  gap: 4px;
  margin-top: 6px;
  width: 100%;
  max-width: 100%;
  overflow-x: hidden;
}
#briefing-content .brief-entry-box,
#briefing-content .brief-tank,
#briefing-content .brief-total-box {
  background: linear-gradient(180deg, #121212 0%, #191919 100%);
  border: 1px solid #6d6d6d;
  border-radius: 8px;
  padding: 5px 5px;
  box-shadow: 0 4px 8px rgba(0,0,0,0.18);
  min-width: 0;
}
#briefing-content .brief-tank.group-start,
#briefing-content .brief-total-box.group-start {
  border-left: 3px solid #8ecf90;
}
#briefing-content .brief-entry-box label,
#briefing-content .brief-tank label,
#briefing-content .brief-total-box label { color: #bbbbbb; margin-bottom: 2px; }
#briefing-content .brief-entry-box input,
#briefing-content .brief-tank input,
#briefing-content .brief-total-box b { color: #fff; font-size: 1.23rem; text-align: center; display: block; }
#briefing-content .leg-inputs .fieldbox input[readonly] { pointer-events: none; }

#briefing-content .fuel-airframe-wrap { margin-top: 0; background: transparent; border: none; padding: 0; }
#briefing-content .fuel-airframe-title { display: none; }
#briefing-content .fuel-wing-row {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  gap: 4px;
  align-items: stretch;
  width: 100%;
  max-width: 100%;
  overflow-x: hidden;
}
#briefing-content .stats-row > *,
#briefing-content .fuel-wing-row > *,
#briefing-content .brief-entry-grid > *,
#briefing-content .leg-grid > * { min-width: 0; }
#briefing-content .brief-tank { min-height: 64px; display: flex; flex-direction: column; justify-content: space-between; }
#briefing-content .brief-tank b,
#briefing-content .brief-tank small,
#briefing-content .fuel-plane-icon,
#briefing-content .brief-fuel-summary { display: none; }
#briefing-content .brief-total-box { min-height: 64px; }
#briefing-content .brief-total-box.warn { border-color: #D32F2F; }
#briefing-content .brief-total-box b { line-height: 1.35; }
#briefing-content .brief-note { margin-top: 6px; color: #e8edf3; font-size: 0.93rem; text-transform: uppercase; }
#briefing-content .brief-warning-inline { margin-top: 4px; color: #ffb3b3; font-size: 0.93rem; text-transform: uppercase; display: none; }

#brief-keypad-modal {
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.58);
  z-index: 10000;
  display: none;
  align-items: center;
  justify-content: center;
  padding: 12px;
}
#brief-keypad-sheet {
  width: min(420px, 96vw);
  background: #111821;
  border: 1px solid #3d4f66;
  border-radius: 12px;
  padding: 10px;
}
#brief-keypad-title { color: #d2dfef; font-size: 1.2rem; margin-bottom: 6px; }
#brief-keypad-display {
  width: 100%;
  height: 48px;
  border: 1px solid #496488;
  border-radius: 8px;
  color: #fff;
  background: #0d141c;
  font-size: 2rem;
  text-align: right;
  padding: 8px 10px;
  margin-bottom: 8px;
}
#brief-keypad-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 6px;
}
#brief-keypad-grid button,
#brief-keypad-actions button {
  height: 46px;
  border-radius: 8px;
  border: 1px solid #4f6380;
  background: #1d2936;
  color: #fff;
  font-size: 1.65rem;
  font-weight: 800;
}
#brief-keypad-actions {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 6px;
  margin-top: 6px;
}
#brief-keypad-actions button.ok { background: #1f6cd1; border-color: #3176d4; }

@media (max-width: 980px) {
  #briefing-content .stats-row { grid-template-columns: repeat(5, 1fr); }
  #briefing-content .fuel-wing-row { grid-template-columns: repeat(7, 1fr); }
  #briefing-content .brief-entry-grid { grid-template-columns: repeat(8, 1fr); }
}
@media (max-width: 768px) {
  #briefing-content { padding: 0 6px 100px 6px; }
  #briefing-content .brief-title { font-size: 1.83rem; }
}
@media (max-width: 720px) {
  #briefing-content .leg-inputs { flex-wrap: wrap; }
  #briefing-content .fieldbox { border-right: none; border-bottom: 1px solid #d0d0d0; }
  #briefing-content .fieldbox:last-child { border-bottom: none; }
}
`;
  (document.head || document.documentElement).appendChild(s);
})();

