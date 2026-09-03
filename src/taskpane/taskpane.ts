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
import { initializeIconPicker } from "./iconPicker";

function disableAutomaticTaskPane(): void {
  Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
  Office.context.document.settings.saveAsync();
}

Office.onReady(() => {
  disableAutomaticTaskPane();
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