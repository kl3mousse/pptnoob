import { withSelectedShapes } from "../office/selection";

function setSelectedShapeMargins(points: number, commandEvent: Office.AddinCommands.Event): Promise<void> {
  return withSelectedShapes(commandEvent, "items", (shapes) => {
    for (const shape of shapes.items) {
      shape.textFrame.leftMargin = points;
      shape.textFrame.rightMargin = points;
      shape.textFrame.topMargin = points;
      shape.textFrame.bottomMargin = points;
    }
  });
}

export function setMarginsPointOneCm(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapeMargins(0.1 * 72 / 2.54, commandEvent);
}

export function setMarginsZero(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapeMargins(0, commandEvent);
}