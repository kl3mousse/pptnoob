import { Favorite } from "./favorite";
import { FavoritesStore } from "./favorites-store";
import { deriveFavoriteName, normalizeFavoriteUrl } from "./favorite-utils";

export interface FavoriteChanges {
  name: string;
  comment?: string;
  url: string;
}

function createFavoriteId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

export class FavoritesService {
  constructor(
    private readonly store: FavoritesStore,
    private readonly createId: () => string = createFavoriteId,
    private readonly now: () => Date = () => new Date(),
  ) {}

  list(): Promise<Favorite[]> {
    return this.store.list();
  }

  async findByUrl(url: string): Promise<Favorite | undefined> {
    const normalizedUrl = normalizeFavoriteUrl(url);
    return (await this.list()).find((favorite) => normalizeFavoriteUrl(favorite.url) === normalizedUrl);
  }

  async add(url: string, name = deriveFavoriteName(url), comment?: string): Promise<Favorite> {
    const existing = await this.findByUrl(url);
    if (existing) return existing;
    const favorite: Favorite = {
      id: this.createId(),
      name: name.trim() || deriveFavoriteName(url),
      url: url.trim(),
      addedAt: this.now().toISOString(),
    };
    if (comment?.trim()) favorite.comment = comment.trim();
    await this.store.add(favorite);
    return favorite;
  }

  async update(id: string, changes: FavoriteChanges): Promise<Favorite> {
    const favorites = await this.list();
    const favorite = favorites.find((item) => item.id === id);
    if (!favorite) throw new Error("Favorite no longer exists.");
    const duplicate = favorites.find((item) => item.id !== id
      && normalizeFavoriteUrl(item.url) === normalizeFavoriteUrl(changes.url));
    if (duplicate) throw new Error("That presentation is already in Favorites.");
    const updated: Favorite = {
      ...favorite,
      name: changes.name.trim() || deriveFavoriteName(changes.url),
      url: changes.url.trim(),
      updatedAt: this.now().toISOString(),
    };
    if (changes.comment?.trim()) updated.comment = changes.comment.trim();
    else delete updated.comment;
    await this.store.update(updated);
    return updated;
  }

  remove(id: string): Promise<void> {
    return this.store.remove(id);
  }

  reorder(ids: string[]): Promise<void> {
    return this.store.reorder(ids);
  }

  async move(id: string, offset: -1 | 1): Promise<void> {
    const favorites = await this.list();
    const index = favorites.findIndex((favorite) => favorite.id === id);
    const destination = index + offset;
    if (index < 0 || destination < 0 || destination >= favorites.length) return;
    [favorites[index], favorites[destination]] = [favorites[destination], favorites[index]];
    await this.store.reorder(favorites.map((favorite) => favorite.id));
  }
}