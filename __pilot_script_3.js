
// Move keypad modal to document.body immediately so container.innerHTML wipes (spinner, re-render)
// can never destroy it. Runs synchronously at parse time, before any tab navigation code.
(function() {
  var m = document.getElementById('brief-keypad-modal');
  if (m && m.parentElement !== document.body) document.body.appendChild(m);
})();

