import { Favorite } from "./favorite";

export interface FavoritesStore {
  list(): Promise<Favorite[]>;
  add(favorite: Favorite): Promise<void>;
  update(favorite: Favorite): Promise<void>;
  remove(id: string): Promise<void>;
  reorder(ids: string[]): Promise<void>;
}