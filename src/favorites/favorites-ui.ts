import { getCurrentPresentationUrl, openContainingFolderUrl, openPresentationUrl, requireWebUrl } from "../office/files";
import { setStatus, showStatusToast } from "../ui/status";
import { Favorite } from "./favorite";
import { deriveFavoriteLocation, deriveFavoriteName, filterFavorites, normalizeFavoriteUrl } from "./favorite-utils";
import { FavoritesService } from "./favorites-service";
import { LocalStorageFavoritesStore } from "./local-storage-favorites-store";

const service = new FavoritesService(new LocalStorageFavoritesStore());
let favorites: Favorite[] = [];
let currentPresentationUrl = "";
let editingId: string | undefined;
let pendingFavorite: Favorite | undefined;
let removingId: string | undefined;

function element<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

function reportError(error: unknown): void {
  setStatus(error instanceof Error ? error.message : String(error));
}

function run(action: () => Promise<void>): void {
  action().catch(reportError);
}

function createButton(label: string, action: () => void, className = "btn btn-ghost btn-sm"): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = className;
  button.textContent = label;
  button.addEventListener("click", action);
  return button;
}

function updateCurrentState(): void {
  const current = currentPresentationUrl
    ? favorites.find((favorite) => normalizeFavoriteUrl(favorite.url) === normalizeFavoriteUrl(currentPresentationUrl))
    : undefined;
  const button = element<HTMLButtonElement>("addCurrentFavorite");
  button.textContent = current ? "★ Current deck saved" : "☆ Add current deck";
  button.classList.toggle("is-favorited", Boolean(current));
  button.dataset.favoriteId = current?.id || "";
}

function createField(labelText: string, control: HTMLElement): HTMLLabelElement {
  const label = document.createElement("label");
  const text = document.createElement("span");
  label.className = "favorite-field";
  text.textContent = labelText;
  label.append(text, control);
  return label;
}

function createEditForm(favorite: Favorite): HTMLElement {
  const isNew = !favorite.id;
  const form = document.createElement("form");
  form.className = "favorite-edit";
  const title = document.createElement("h3");
  title.textContent = isNew ? "Add to favorites" : "Edit favorite";
  const name = document.createElement("input");
  name.className = "control control-compact";
  name.value = favorite.name;
  name.required = true;
  const note = document.createElement("input");
  note.className = "control control-compact";
  note.value = favorite.comment || "";
  note.placeholder = "Optional note";
  note.maxLength = 120;
  const url = document.createElement("input");
  url.className = "control control-compact";
  url.type = "url";
  url.value = favorite.url;
  url.required = true;
  const actions = document.createElement("div");
  actions.className = "favorite-form-actions";
  const cancel = createButton("Cancel", () => {
    editingId = undefined;
    pendingFavorite = undefined;
    removingId = undefined;
    renderFavorites();
  });
  const submit = createButton(isNew ? "Add" : "Save", () => undefined, "btn btn-primary btn-sm");
  submit.type = "submit";
  actions.append(cancel, submit);
  form.append(title, createField("Name", name), createField("Note", note));
  if (isNew) {
    const location = document.createElement("div");
    location.className = "favorite-form-location";
    location.textContent = deriveFavoriteLocation(favorite.url);
    location.title = favorite.url;
    form.append(createField("Location", location));
  } else {
    form.append(createField("Location", url));
  }
  form.append(actions);
  if (!isNew) {
    const removeArea = document.createElement("div");
    removeArea.className = "favorite-remove-area";
    if (removingId === favorite.id) {
      const prompt = document.createElement("span");
      prompt.textContent = "Remove this favorite?";
      removeArea.append(
        prompt,
        createButton("Cancel", () => {
          removingId = undefined;
          renderFavorites();
        }),
        createButton("Remove", () => run(async () => {
          await service.remove(favorite.id);
          editingId = undefined;
          removingId = undefined;
          await refreshFavorites();
          showStatusToast("✓ Removed from favorites");
        }), "btn btn-danger btn-sm"),
      );
    } else {
      removeArea.append(createButton("Remove favorite", () => {
        removingId = favorite.id;
        renderFavorites();
      }, "btn btn-link btn-danger-text btn-sm"));
    }
    form.append(removeArea);
  }
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    run(async () => {
      requireWebUrl(url.value);
      if (favorite.id) {
        await service.update(favorite.id, { name: name.value, comment: note.value, url: url.value });
      } else {
        await service.add(url.value, name.value, note.value);
      }
      editingId = undefined;
      pendingFavorite = undefined;
      await refreshFavorites();
      showStatusToast(favorite.id ? "✓ Favorite updated" : "✓ Added to favorites");
    });
  });
  queueMicrotask(() => name.focus());
  return form;
}

function createFavoriteCard(favorite: Favorite): HTMLElement {
  const card = document.createElement("article");
  card.className = "favorite-card";
  card.dataset.favoriteId = favorite.id;
  if (editingId === favorite.id) {
    card.append(createEditForm(favorite));
    return card;
  }

  const main = document.createElement("button");
  main.type = "button";
  main.className = "favorite-main";
  main.title = `Open ${favorite.name}\n${favorite.url}`;
  main.setAttribute("aria-label", `Open ${favorite.name}`);
  if (favorite.comment) {
    const note = document.createElement("span");
    note.className = "favorite-note";
    note.textContent = favorite.comment;
    main.append(note);
  }
  const name = document.createElement("strong");
  name.textContent = favorite.name;
  const location = document.createElement("span");
  location.className = "favorite-location";
  location.textContent = deriveFavoriteLocation(favorite.url);
  location.title = favorite.url;
  main.append(name, location);
  main.addEventListener("click", () => {
    try {
      openPresentationUrl(favorite.url);
      showStatusToast("Opening in PowerPoint…");
    } catch (error) {
      reportError(error);
    }
  });
  const menu = document.createElement("details");
  menu.className = "action-menu";
  const menuTrigger = document.createElement("summary");
  menuTrigger.className = "icon-action";
  menuTrigger.textContent = "⋯";
  menuTrigger.title = `More actions for ${favorite.name}`;
  menuTrigger.setAttribute("role", "button");
  menuTrigger.setAttribute("aria-label", `More actions for ${favorite.name}`);
  menuTrigger.setAttribute("aria-haspopup", "menu");
  menuTrigger.setAttribute("aria-expanded", "false");
  const menuItems = document.createElement("div");
  menuItems.className = "action-menu-items";
  menuItems.setAttribute("role", "menu");
  const createMenuItem = (label: string, action: () => void, disabled = false): HTMLButtonElement => {
    const item = createButton(label, () => {
      menu.open = false;
      action();
    }, "action-menu-item");
    item.disabled = disabled;
    item.setAttribute("role", "menuitem");
    return item;
  };
  const index = favorites.findIndex((item) => item.id === favorite.id);
  menuItems.append(
    createMenuItem("Open containing folder", () => {
      try {
        openContainingFolderUrl(favorite.url);
        showStatusToast("✓ Opened containing folder");
      } catch (error) {
        reportError(error);
      }
    }),
    createMenuItem("Edit", () => {
      editingId = favorite.id;
      removingId = undefined;
      renderFavorites();
    }),
    createMenuItem("Move up", () => run(async () => {
      await service.move(favorite.id, -1);
      await refreshFavorites();
    }), index === 0),
    createMenuItem("Move down", () => run(async () => {
      await service.move(favorite.id, 1);
      await refreshFavorites();
    }), index === favorites.length - 1),
    createMenuItem("Remove", () => {
      editingId = favorite.id;
      removingId = favorite.id;
      renderFavorites();
    }),
  );
  menu.append(menuTrigger, menuItems);
  menu.addEventListener("toggle", () => {
    menuTrigger.setAttribute("aria-expanded", String(menu.open));
    if (!menu.open) return;
    document.querySelectorAll<HTMLDetailsElement>(".action-menu[open]").forEach((other) => {
      if (other !== menu) other.open = false;
    });
  });
  card.append(main, menu);
  return card;
}

function getFilteredFavorites(): Favorite[] {
  return filterFavorites(favorites, element<HTMLInputElement>("favoriteSearch").value);
}

function renderFavorites(): void {
  const matches = getFilteredFavorites();
  const cards = matches.map(createFavoriteCard);
  if (pendingFavorite) cards.unshift(createFavoriteCard(pendingFavorite));
  element("favoritesList").replaceChildren(...cards);
  const empty = element("emptyFavorites");
  empty.hidden = matches.length > 0 || Boolean(pendingFavorite);
  empty.textContent = favorites.length ? "No favorites match your search." : "No favorites yet.";
  updateCurrentState();
}

async function refreshFavorites(): Promise<void> {
  favorites = await service.list();
  renderFavorites();
}

async function addCurrentPresentation(): Promise<void> {
  if (!currentPresentationUrl) {
    throw new Error("Save this presentation to SharePoint, OneDrive, or another supported web location before adding it.");
  }
  requireWebUrl(currentPresentationUrl);
  const existing = await service.findByUrl(currentPresentationUrl);
  if (existing) {
    editingId = existing.id;
    pendingFavorite = undefined;
    removingId = undefined;
    renderFavorites();
    queueMicrotask(() => {
      const card = document.querySelector<HTMLElement>(`[data-favorite-id="${existing.id}"]`);
      card?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    });
    return;
  }
  pendingFavorite = {
    id: "",
    name: deriveFavoriteName(currentPresentationUrl),
    url: currentPresentationUrl,
    addedAt: "",
  };
  editingId = "";
  renderFavorites();
}

async function loadCurrentPresentation(): Promise<void> {
  try {
    currentPresentationUrl = await getCurrentPresentationUrl();
  } catch (error) {
    currentPresentationUrl = "";
    reportError(error);
  }
  updateCurrentState();
}

export function initializeFavorites(): void {
  element("addCurrentFavorite").addEventListener("click", () => run(addCurrentPresentation));
  element("favoriteSearch").addEventListener("input", renderFavorites);
  element("favoriteSearch").addEventListener("keydown", (event) => {
    if (event.key !== "Enter") return;
    const first = getFilteredFavorites()[0];
    if (!first) return;
    event.preventDefault();
    try {
      openPresentationUrl(first.url);
      showStatusToast("Opening in PowerPoint…");
    } catch (error) {
      reportError(error);
    }
  });
  document.addEventListener("click", (event) => {
    const target = event.target;
    if (target instanceof Element && target.closest(".action-menu")) return;
    document.querySelectorAll<HTMLDetailsElement>(".action-menu[open]").forEach((menu) => {
      menu.open = false;
    });
  });
  run(async () => {
    await Promise.all([refreshFavorites(), loadCurrentPresentation()]);
  });
}