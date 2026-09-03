import { setStatus } from "../ui/status";
import { getCurrentPresentationUrl, openContainingFolderUrl } from "../office/files";

export async function openContainingFolder(commandEvent?: Office.AddinCommands.Event): Promise<void> {
  try {
    const fileUrl = await getCurrentPresentationUrl();
    if (!fileUrl) {
      throw new Error("Save the presentation before opening its folder.");
    }
    openContainingFolderUrl(fileUrl);
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