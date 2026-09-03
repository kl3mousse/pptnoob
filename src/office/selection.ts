type SelectedShapesHandler = (
  shapes: PowerPoint.ShapeScopedCollection,
  context: PowerPoint.RequestContext,
) => void | Promise<void>;

export async function withSelectedShapes(
  commandEvent: Office.AddinCommands.Event,
  properties: string,
  handler: SelectedShapesHandler,
): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load(properties);
      await context.sync();
      await handler(shapes, context);
      await context.sync();
    });
  } finally {
    commandEvent.completed();
  }
}