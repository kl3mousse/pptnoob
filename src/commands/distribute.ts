import { withSelectedShapes } from "../office/selection";
import { getSelectionBounds, sortLeftToRight, sortTopToBottom } from "../office/shapes";

export function distributeHorizontally(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return withSelectedShapes(commandEvent, "items/left,items/top,items/width,items/height", (shapes) => {
    if (shapes.items.length < 3) return;

    const bounds = getSelectionBounds(shapes.items);
    const orderedShapes = sortLeftToRight(shapes.items);
    const totalWidth = orderedShapes.reduce((sum, shape) => sum + shape.width, 0);
    const gap = (bounds.right - bounds.left - totalWidth) / (orderedShapes.length - 1);
    let nextLeft = bounds.left;

    orderedShapes.forEach((shape) => {
      shape.left = nextLeft;
      nextLeft += shape.width + gap;
    });
  });
}

export function distributeVertically(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return withSelectedShapes(commandEvent, "items/left,items/top,items/width,items/height", (shapes) => {
    if (shapes.items.length < 3) return;

    const bounds = getSelectionBounds(shapes.items);
    const orderedShapes = sortTopToBottom(shapes.items);
    const totalHeight = orderedShapes.reduce((sum, shape) => sum + shape.height, 0);
    const gap = (bounds.bottom - bounds.top - totalHeight) / (orderedShapes.length - 1);
    let nextTop = bounds.top;

    orderedShapes.forEach((shape) => {
      shape.top = nextTop;
      nextTop += shape.height + gap;
    });
  });
}