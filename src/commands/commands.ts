/* global Office */

// The ribbon button opens the task pane directly (no function command), so there is
// nothing to register here. This file only exists so commands.html (the classic XML
// manifest's FunctionFile) loads Office.js and stays a valid, empty command surface.
Office.onReady(() => {
  // Intentionally empty.
});
