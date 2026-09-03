import {
  IconDefinition,
  insertPhosphorIcon,
  loadPhosphorIcons,
} from "../icons/phosphor";
import { setStatus, showStatusToast } from "../ui/status";

const pageSize = 60;

let icons: IconDefinition[] = [];
let visibleIconCount = pageSize;

function getMatchingIcons(): IconDefinition[] {
  const query = (document.getElementById("iconSearch") as HTMLInputElement).value.trim().toLowerCase();
  const weight = (document.getElementById("iconWeight") as HTMLSelectElement).value;
  return icons.filter((icon) => icon.weight === weight && icon.slug.includes(query));
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

function resetIconResults(): void {
  visibleIconCount = pageSize;
  renderIcons();
}

function showMoreIcons(): void {
  visibleIconCount += pageSize;
  renderIcons();
}

async function insertIcon(icon: IconDefinition): Promise<void> {
  try {
    setStatus(`Inserting ${icon.name}...`);
    const color = (document.getElementById("iconColor") as HTMLInputElement).value;
    await insertPhosphorIcon(icon, color);
    showStatusToast(`✓ Inserted ${icon.name}`);
  } catch (error: unknown) {
    setStatus(error instanceof Error ? error.message : String(error));
  }
}

async function loadIcons(): Promise<void> {
  try {
    setStatus("Loading Phosphor icons...");
    const library = await loadPhosphorIcons();
    icons = library.icons;
    renderIcons();
    showStatusToast(`${library.iconsPerWeight.toLocaleString()} Phosphor icons available`);
  } catch (error: unknown) {
    setStatus(error instanceof Error ? error.message : String(error));
  }
}

export function initializeIconPicker(): void {
  document.getElementById("iconSearch")!.addEventListener("input", resetIconResults);
  document.getElementById("iconWeight")!.addEventListener("change", resetIconResults);
  document.getElementById("showMoreIcons")!.addEventListener("click", showMoreIcons);
  loadIcons();
}