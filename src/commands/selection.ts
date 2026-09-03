import { withSelectedShapes } from "../office/selection";

function setSelectedShapesZOrder(
  position: PowerPoint.ShapeZOrder,
  commandEvent: Office.AddinCommands.Event,
): Promise<void> {
  return withSelectedShapes(commandEvent, "items/zOrderPosition", (shapes) => {
    const orderedShapes = [...shapes.items].sort((first, second) =>
      position === PowerPoint.ShapeZOrder.bringToFront
        ? first.zOrderPosition - second.zOrderPosition
        : second.zOrderPosition - first.zOrderPosition,
    );
    orderedShapes.forEach((shape) => shape.setZOrder(position));
  });
}

export function bringToFront(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapesZOrder(PowerPoint.ShapeZOrder.bringToFront, commandEvent);
}

export function sendToBack(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapesZOrder(PowerPoint.ShapeZOrder.sendToBack, commandEvent);
}