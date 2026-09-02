/// <reference types="office-js" />

interface IconDefinition {
  name: string;
  slug: string;
  weight: string;
  path: string;
}

type IconIndex = Record<string, string[]>;

const pageSize = 60;
let icons: IconDefinition[] = [];
let visibleIconCount = pageSize;

Office.onReady(() => {
  disableAutomaticTaskPane();
  document.getElementById("iconSearch")!.addEventListener("input", resetIconResults);
  document.getElementById("iconWeight")!.addEventListener("change", resetIconResults);
  document.getElementById("showMoreIcons")!.addEventListener("click", showMoreIcons);
  loadIcons();
});

function disableAutomaticTaskPane(): void {
  Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
  Office.context.document.settings.saveAsync();
}

function setStatus(message: string): void {
  const status = document.getElementById("status");
  if (status) status.textContent = message;
}

async function loadIcons(): Promise<void> {
  try {
    setStatus("Loading Phosphor icons...");
    const response = await fetch("assets/phosphor/icon-index.json");
    if (!response.ok) throw new Error("Could not load the Phosphor icon index.");
    const index = await response.json() as IconIndex;
    icons = Object.entries(index).flatMap(([weight, filenames]) => filenames.map((filename) => {
      const suffix = weight === "regular" ? "" : `-${weight}`;
      const slug = filename.slice(0, -4).replace(new RegExp(`${suffix}$`), "");
      return {
        name: slug.split("-").map((word) => word.charAt(0).toUpperCase() + word.slice(1)).join(" "),
        slug,
        weight,
        path: `assets/phosphor/phosphor-icons/SVGs/${weight}/${filename}`,
      };
    }));
    renderIcons();
    const weightCount = Object.keys(index).length;
    const iconsPerWeight = Object.values(index)[0]?.length ?? 0;
    setStatus(`${iconsPerWeight.toLocaleString()} Phosphor icons available in ${weightCount} weights.`);
  } catch (error: unknown) {
    setStatus(error instanceof Error ? error.message : String(error));
  }
}

function resetIconResults(): void {
  visibleIconCount = pageSize;
  renderIcons();
}

function getMatchingIcons(): IconDefinition[] {
  const query = (document.getElementById("iconSearch") as HTMLInputElement).value.trim().toLowerCase();
  const weight = (document.getElementById("iconWeight") as HTMLSelectElement).value;
  return icons.filter((icon) => icon.weight === weight && icon.slug.includes(query));
}

function renderIcons(): void {
  const matches = getMatchingIcons();
  const grid = document.getElementById("iconGrid")!;
  const empty = document.getElementById("emptyIcons")!;
  const showMore = document.getElementById("showMoreIcons") as HTMLButtonElement;
  const resultCount = document.getElementById("iconResultCount")!;

  grid.replaceChildren(...matches.slice(0, visibleIconCount).map(createIconButton));
  empty.hidden = matches.length > 0;
  showMore.hidden = matches.length <= visibleIconCount;
  resultCount.textContent = `${Math.min(matches.length, visibleIconCount).toLocaleString()} of ${matches.length.toLocaleString()}`;
}

function showMoreIcons(): void {
  visibleIconCount += pageSize;
  renderIcons();
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

