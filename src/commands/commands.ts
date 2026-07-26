/**
 * No ExecuteFunction ribbon actions are defined in this add-in (the ribbon
 * button uses ShowTaskpane instead), so this file only needs to satisfy
 * Office.js's initialization contract by resolving onReady. There is
 * nothing to register with Office.actions.associate() because we have no
 * ExecuteFunction-triggered commands.
 */
Office.onReady(() => {
  // Intentionally empty: this function file exists only because the
  // manifest's DesktopFormFactor requires one. See README for details.
});
