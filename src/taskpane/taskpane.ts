/// <reference types="office-js" />

Office.onReady(() => {
  document.getElementById("openFolderBtn")!.addEventListener("click", () => openContainingFolder());
});

function getFileUrl(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFilePropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.url);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

async function openContainingFolder(commandEvent?: Office.AddinCommands.Event): Promise<void> {
  const status = document.getElementById("status");

  try {
    const fileUrl = await getFileUrl();
    if (!fileUrl) {
      throw new Error("Save the presentation before opening its folder.");
    }

    const url = new URL(fileUrl);
    if (url.protocol !== "https:" && url.protocol !== "http:") {
      throw new Error("Finder cannot be opened from an Office web add-in. This command supports SharePoint and OneDrive files.");
    }

    url.pathname = url.pathname.substring(0, url.pathname.lastIndexOf("/") + 1);
    url.search = "";
    url.hash = "";
    Office.context.ui.openBrowserWindow(url.toString());
    if (status) status.textContent = "Opened the containing folder in your browser.";
  } catch (error: unknown) {
    if (status) status.textContent = error instanceof Error ? error.message : String(error);
  } finally {
    commandEvent?.completed();
  }
}

(globalThis as typeof globalThis & { openContainingFolder: typeof openContainingFolder }).openContainingFolder = openContainingFolder;

async function setSelectedShapeMargins(points: number, commandEvent: Office.AddinCommands.Event): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items");
      await context.sync();

      for (const shape of shapes.items) {
        shape.textFrame.leftMargin = points;
        shape.textFrame.rightMargin = points;
        shape.textFrame.topMargin = points;
        shape.textFrame.bottomMargin = points;
      }
      await context.sync();
    });
  } finally {
    commandEvent.completed();
  }
}

function setMarginsPointOneCm(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapeMargins(0.1 * 72 / 2.54, commandEvent);
}

function setMarginsZero(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapeMargins(0, commandEvent);
}

Object.assign(globalThis, { setMarginsPointOneCm, setMarginsZero });

