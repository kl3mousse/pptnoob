import { withSelectedShapes } from "../office/selection";
import { getSelectionBounds } from "../office/shapes";

type Alignment = "left" | "center" | "right" | "top" | "middle" | "bottom";

function alignSelectedShapes(alignment: Alignment, commandEvent: Office.AddinCommands.Event): Promise<void> {
  return withSelectedShapes(commandEvent, "items/left,items/top,items/width,items/height", (shapes) => {
    if (shapes.items.length < 2) return;

    const bounds = getSelectionBounds(shapes.items);

    if (alignment === "left") {
      shapes.items.forEach((shape) => { shape.left = bounds.left; });
    } else if (alignment === "center") {
      const center = (bounds.left + bounds.right) / 2;
      shapes.items.forEach((shape) => { shape.left = center - shape.width / 2; });
    } else if (alignment === "right") {
      shapes.items.forEach((shape) => { shape.left = bounds.right - shape.width; });
    } else if (alignment === "top") {
      shapes.items.forEach((shape) => { shape.top = bounds.top; });
    } else if (alignment === "middle") {
      const middle = (bounds.top + bounds.bottom) / 2;
      shapes.items.forEach((shape) => { shape.top = middle - shape.height / 2; });
    } else {
      shapes.items.forEach((shape) => { shape.top = bounds.bottom - shape.height; });
    }
  });
}

export function alignLeft(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return alignSelectedShapes("left", commandEvent);
}

export function alignCenter(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return alignSelectedShapes("center", commandEvent);
}

export function alignRight(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return alignSelectedShapes("right", commandEvent);
}

export function alignTop(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return alignSelectedShapes("top", commandEvent);
}

export function alignMiddle(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return alignSelectedShapes("middle", commandEvent);
}

export function alignBottom(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return alignSelectedShapes("bottom", commandEvent);
}