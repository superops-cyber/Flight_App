

function renderTankSection(mCap, tCap, launchFuel) {
  function tank(key, fullLabel, cap, type) {
    return `
      <div class="brief-tank ${type}">
        <label>${key}</label>
        <input type="number"
               class="brief-tank-input"
               inputmode="numeric"
               data-tank-key="${key}"
               data-max="${cap}"
               min="0"
               step="1"
               oninput="calculateBriefFuelTally()">
      </div>`;
  }

  const safeMainCap = Math.max(0, Number(mCap || 0));
  const safeTipCap = Math.max(0, Number(tCap || 0));

  return `
  <div class="fuel-airframe-wrap">
    <div class="fuel-airframe-title">Fuel On Board by Tank (Cessna Layout)</div>
    <div class="fuel-wing-row">
      ${tank("LT", "Left Tip (LT)", safeTipCap, "tip")}
      ${tank("LM", "Left Main (LM)", safeMainCap, "main")}
      ${tank("RM", "Right Main (RM)", safeMainCap, "main")}
      ${tank("RT", "Right Tip (RT)", safeTipCap, "tip")}
      <div id="brief-fuel-total-box" class="brief-total-box">
        <label>Total</label>
        <b id="brief-fuel-tally" data-launch="${launchFuel}">0L</b>
      </div>
    </div>
    <div id="brief-fuel-warning" class="brief-warning-inline">⚠ Fuel mismatch: 0L below planned</div>
  </div>`;
}


