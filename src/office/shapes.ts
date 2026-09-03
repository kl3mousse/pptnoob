export interface ShapeBounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SelectionBounds {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

export function getSelectionBounds(shapes: readonly ShapeBounds[]): SelectionBounds {
  return {
    left: Math.min(...shapes.map((shape) => shape.left)),
    right: Math.max(...shapes.map((shape) => shape.left + shape.width)),
    top: Math.min(...shapes.map((shape) => shape.top)),
    bottom: Math.max(...shapes.map((shape) => shape.top + shape.height)),
  };
}

export function sortLeftToRight<T extends Pick<ShapeBounds, "left">>(shapes: readonly T[]): T[] {
  return [...shapes].sort((first, second) => first.left - second.left);
}

export function sortTopToBottom<T extends Pick<ShapeBounds, "top">>(shapes: readonly T[]): T[] {
  return [...shapes].sort((first, second) => first.top - second.top);
}