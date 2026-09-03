/// <reference types="office-js" />

import {
  alignBottom,
  alignCenter,
  alignLeft,
  alignMiddle,
  alignRight,
  alignTop,
} from "../commands/align";
import { distributeHorizontally, distributeVertically } from "../commands/distribute";
import { openContainingFolder, openInfo } from "../commands/file";
import { setMarginsPointOneCm, setMarginsZero } from "../commands/margins";
import { bringToFront, sendToBack } from "../commands/selection";
import { initializeFavorites } from "../favorites/favorites-ui";
import { initializeIconPicker } from "./iconPicker";

type PaneName = "favorites" | "icons";

function disableAutomaticTaskPane(): void {
  Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
  Office.context.document.settings.saveAsync();
}

function showPane(paneName: PaneName): void {
  document.querySelectorAll<HTMLElement>("[data-pane]").forEach((pane) => {
    pane.hidden = pane.dataset.pane !== paneName;
  });
  document.querySelectorAll<HTMLButtonElement>("[data-pane-target]").forEach((tab) => {
    const active = tab.dataset.paneTarget === paneName;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-selected", String(active));
  });
}

function initializePaneNavigation(): void {
  document.querySelectorAll<HTMLButtonElement>("[data-pane-target]").forEach((tab) => {
    tab.addEventListener("click", () => {
      const paneName = tab.dataset.paneTarget as PaneName;
      showPane(paneName);
      window.history.replaceState(null, "", `#${paneName}`);
    });
  });
  showPane(window.location.hash === "#icons" ? "icons" : "favorites");
}

Office.onReady(() => {
  disableAutomaticTaskPane();
  initializePaneNavigation();
  initializeFavorites();
  initializeIconPicker();
});

Object.assign(globalThis, {
  alignBottom,
  alignCenter,
  alignLeft,
  alignMiddle,
  alignRight,
  alignTop,
  bringToFront,
  distributeHorizontally,
  distributeVertically,
  openContainingFolder,
  openInfo,
  sendToBack,
  setMarginsPointOneCm,
  setMarginsZero,
});