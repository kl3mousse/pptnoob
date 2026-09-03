import { setStatus } from "../ui/status";

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

export async function openContainingFolder(commandEvent?: Office.AddinCommands.Event): Promise<void> {
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
    setStatus("Opened the containing folder in your browser.");
  } catch (error: unknown) {
    setStatus(error instanceof Error ? error.message : String(error));
  } finally {
    commandEvent?.completed();
  }
}

export function openInfo(commandEvent: Office.AddinCommands.Event): void {
  const infoUrl = new URL("info.html", window.location.href).toString();
  Office.context.ui.displayDialogAsync(infoUrl, { height: 55, width: 35 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      setStatus(`Could not open add-in information: ${result.error.message}`);
    }
    commandEvent.completed();
  });
}