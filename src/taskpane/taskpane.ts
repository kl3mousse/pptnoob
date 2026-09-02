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
    icons = Object.entries(index).reduce<IconDefinition[]>((allIcons, [weight, filenames]) => allIcons.concat(filenames.map((filename) => {
      const suffix = weight === "regular" ? "" : `-${weight}`;
      const slug = filename.slice(0, -4).replace(new RegExp(`${suffix}$`), "");
      return {
        name: slug.split("-").map((word) => word.charAt(0).toUpperCase() + word.slice(1)).join(" "),
        slug,
        weight,
        path: `assets/phosphor/phosphor-icons/SVGs/${weight}/${filename}`,
      };
    })), []);
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

type ShapeLayoutCommand =
  | "alignLeft"
  | "alignCenter"
  | "alignRight"
  | "alignTop"
  | "alignMiddle"
  | "alignBottom"
  | "distributeHorizontally"
  | "distributeVertically";

async function layoutSelectedShapes(command: ShapeLayoutCommand, commandEvent: Office.AddinCommands.Event): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/left,items/top,items/width,items/height");
      await context.sync();

      if (shapes.items.length < 2) return;

      const left = Math.min(...shapes.items.map((shape) => shape.left));
      const right = Math.max(...shapes.items.map((shape) => shape.left + shape.width));
      const top = Math.min(...shapes.items.map((shape) => shape.top));
      const bottom = Math.max(...shapes.items.map((shape) => shape.top + shape.height));

      if (command === "alignLeft") {
        shapes.items.forEach((shape) => { shape.left = left; });
      } else if (command === "alignCenter") {
        const center = (left + right) / 2;
        shapes.items.forEach((shape) => { shape.left = center - shape.width / 2; });
      } else if (command === "alignRight") {
        shapes.items.forEach((shape) => { shape.left = right - shape.width; });
      } else if (command === "alignTop") {
        shapes.items.forEach((shape) => { shape.top = top; });
      } else if (command === "alignMiddle") {
        const middle = (top + bottom) / 2;
        shapes.items.forEach((shape) => { shape.top = middle - shape.height / 2; });
      } else if (command === "alignBottom") {
        shapes.items.forEach((shape) => { shape.top = bottom - shape.height; });
      } else if (shapes.items.length >= 3 && command === "distributeHorizontally") {
        const orderedShapes = [...shapes.items].sort((first, second) => first.left - second.left);
        const totalWidth = orderedShapes.reduce((sum, shape) => sum + shape.width, 0);
        const gap = (right - left - totalWidth) / (orderedShapes.length - 1);
        let nextLeft = left;
        orderedShapes.forEach((shape) => {
          shape.left = nextLeft;
          nextLeft += shape.width + gap;
        });
      } else if (shapes.items.length >= 3 && command === "distributeVertically") {
        const orderedShapes = [...shapes.items].sort((first, second) => first.top - second.top);
        const totalHeight = orderedShapes.reduce((sum, shape) => sum + shape.height, 0);
        const gap = (bottom - top - totalHeight) / (orderedShapes.length - 1);
        let nextTop = top;
        orderedShapes.forEach((shape) => {
          shape.top = nextTop;
          nextTop += shape.height + gap;
        });
      }

      await context.sync();
    });
  } finally {
    commandEvent.completed();
  }
}

function alignLeft(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("alignLeft", commandEvent);
}

function alignCenter(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("alignCenter", commandEvent);
}

function alignRight(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("alignRight", commandEvent);
}

function alignTop(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("alignTop", commandEvent);
}

function alignMiddle(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("alignMiddle", commandEvent);
}

function alignBottom(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("alignBottom", commandEvent);
}

function distributeHorizontally(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("distributeHorizontally", commandEvent);
}

function distributeVertically(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return layoutSelectedShapes("distributeVertically", commandEvent);
}

async function setSelectedShapesZOrder(position: PowerPoint.ShapeZOrder, commandEvent: Office.AddinCommands.Event): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/zOrderPosition");
      await context.sync();

      const orderedShapes = [...shapes.items].sort((first, second) =>
        position === PowerPoint.ShapeZOrder.bringToFront
          ? first.zOrderPosition - second.zOrderPosition
          : second.zOrderPosition - first.zOrderPosition,
      );
      orderedShapes.forEach((shape) => shape.setZOrder(position));
      await context.sync();
    });
  } finally {
    commandEvent.completed();
  }
}

function bringToFront(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapesZOrder(PowerPoint.ShapeZOrder.bringToFront, commandEvent);
}

function sendToBack(commandEvent: Office.AddinCommands.Event): Promise<void> {
  return setSelectedShapesZOrder(PowerPoint.ShapeZOrder.sendToBack, commandEvent);
}

function openInfo(commandEvent: Office.AddinCommands.Event): void {
  const infoUrl = new URL("info.html", window.location.href).toString();
  Office.context.ui.displayDialogAsync(infoUrl, { height: 55, width: 35 }, (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      setStatus(`Could not open add-in information: ${result.error.message}`);
    }
    commandEvent.completed();
  });
}

Object.assign(globalThis, {
  alignBottom,
  alignCenter,
  alignLeft,
  alignMiddle,
  alignRight,
  alignTop,
  bringToFront,
  distributeHorizontally,
  distributeVertically,
  openInfo,
  sendToBack,
  setMarginsPointOneCm,
  setMarginsZero,
});

