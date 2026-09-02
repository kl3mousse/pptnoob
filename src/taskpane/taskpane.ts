/// <reference types="office-js" />

interface IconDefinition {
  name: string;
  keywords: string[];
  path: string;
}

const icons: IconDefinition[] = [
  {
    name: "Microsoft PowerPoint",
    keywords: ["microsoft", "office", "powerpoint", "presentation", "slides"],
    path: "assets/phosphor/microsoft-powerpoint-logo.svg",
  },
];

Office.onReady(() => {
  document.getElementById("openFolderBtn")!.addEventListener("click", () => openContainingFolder());
  document.getElementById("iconSearch")!.addEventListener("input", renderIcons);
  renderIcons();
});

function setStatus(message: string): void {
  const status = document.getElementById("status");
  if (status) status.textContent = message;
}

function renderIcons(): void {
  const query = (document.getElementById("iconSearch") as HTMLInputElement).value.trim().toLowerCase();
  const grid = document.getElementById("iconGrid")!;
  const empty = document.getElementById("emptyIcons")!;
  const matches = icons.filter((icon) => [icon.name, ...icon.keywords].some((value) => value.toLowerCase().includes(query)));

  grid.replaceChildren(...matches.map(createIconButton));
  empty.hidden = matches.length > 0;
}

function createIconButton(icon: IconDefinition): HTMLButtonElement {
  const button = document.createElement("button");
  const image = document.createElement("img");
  const label = document.createElement("span");

  button.type = "button";
  button.className = "icon-button";
  button.title = `Insert ${icon.name}`;
  image.src = icon.path;
  image.alt = "";
  label.textContent = icon.name;
  button.append(image, label);
  button.addEventListener("click", () => insertIcon(icon));
  return button;
}

async function insertIcon(icon: IconDefinition): Promise<void> {
  try {
    setStatus(`Inserting ${icon.name}...`);
    const response = await fetch(icon.path);
    if (!response.ok) throw new Error(`Could not load ${icon.name}.`);

    const color = (document.getElementById("iconColor") as HTMLInputElement).value;
    const svg = (await response.text()).replace(/currentColor/g, color);
    await new Promise<void>((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(
        svg,
        { coercionType: Office.CoercionType.XmlSvg },
        (result) => result.status === Office.AsyncResultStatus.Succeeded
          ? resolve()
          : reject(new Error(result.error.message)),
      );
    });
    setStatus(`Inserted ${icon.name}.`);
  } catch (error: unknown) {
    setStatus(error instanceof Error ? error.message : String(error));
  }
}

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

