import { Favorite, isFavorite } from "./favorite";
import { FavoritesStore } from "./favorites-store";

const favoritesStorageKey = "pptnoob.favorites.v1";

interface StorageLike {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export class LocalStorageFavoritesStore implements FavoritesStore {
  constructor(private readonly storage: StorageLike = localStorage) {}

  async list(): Promise<Favorite[]> {
    const value = this.storage.getItem(favoritesStorageKey);
    if (!value) return [];
    try {
      const parsed: unknown = JSON.parse(value);
      return Array.isArray(parsed) ? parsed.filter(isFavorite) : [];
    } catch {
      return [];
    }
  }

  async add(favorite: Favorite): Promise<void> {
    const favorites = await this.list();
    favorites.push(favorite);
    this.save(favorites);
  }

  async update(favorite: Favorite): Promise<void> {
    const favorites = await this.list();
    const index = favorites.findIndex((item) => item.id === favorite.id);
    if (index < 0) throw new Error("Favorite no longer exists.");
    favorites[index] = favorite;
    this.save(favorites);
  }

  async remove(id: string): Promise<void> {
    this.save((await this.list()).filter((favorite) => favorite.id !== id));
  }

  async reorder(ids: string[]): Promise<void> {
    const favorites = await this.list();
    if (ids.length !== favorites.length || new Set(ids).size !== favorites.length) {
      throw new Error("Favorites could not be reordered because the list changed.");
    }
    const byId = new Map(favorites.map((favorite) => [favorite.id, favorite]));
    const reordered = ids.map((id) => byId.get(id));
    if (reordered.some((favorite) => !favorite)) {
      throw new Error("Favorites could not be reordered because the list changed.");
    }
    this.save(reordered as Favorite[]);
  }

  private save(favorites: Favorite[]): void {
    this.storage.setItem(favoritesStorageKey, JSON.stringify(favorites));
  }
}