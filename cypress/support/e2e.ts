// "ResizeObserver loop completed with undelivered notifications" is a benign browser notice fired
// when a ResizeObserver callback (here the toolbar's responsive relayout) triggers another layout
// pass. It is not an application error, so it must not fail a test.
Cypress.on("uncaught:exception", (err) => {
  if (/ResizeObserver loop/.test(err.message)) return false;
  return undefined;
});
