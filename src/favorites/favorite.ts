export interface Favorite {
  id: string;
  name: string;
  comment?: string;
  url: string;
  addedAt: string;
  updatedAt?: string;
  [key: string]: unknown;
}

export function isFavorite(value: unknown): value is Favorite {
  if (!value || typeof value !== "object") return false;
  const favorite = value as Record<string, unknown>;
  return typeof favorite.id === "string"
    && typeof favorite.name === "string"
    && typeof favorite.url === "string"
    && typeof favorite.addedAt === "string"
    && (favorite.comment === undefined || typeof favorite.comment === "string")
    && (favorite.updatedAt === undefined || typeof favorite.updatedAt === "string");
}