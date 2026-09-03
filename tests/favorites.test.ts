/// <reference types="node" />

import assert from "node:assert/strict";
import test from "node:test";
import { Favorite } from "../src/favorites/favorite";
import { deriveFavoriteLocation, deriveFavoriteName, filterFavorites, normalizeFavoriteUrl } from "../src/favorites/favorite-utils";
import { FavoritesService } from "../src/favorites/favorites-service";
import { LocalStorageFavoritesStore } from "../src/favorites/local-storage-favorites-store";
import { getPowerPointOpenUrl } from "../src/office/files";

const storageKey = "pptnoob.favorites.v1";

class MemoryStorage {
  private readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}

function createFavorite(overrides: Partial<Favorite> = {}): Favorite {
  return {
    id: "favorite-1",
    name: "Steering Committee",
    url: "https://example.com/docs/steerco.pptx",
    addedAt: "2026-01-01T00:00:00.000Z",
    ...overrides,
  };
}

test("empty and malformed storage produce an empty list", async () => {
  const storage = new MemoryStorage();
  const store = new LocalStorageFavoritesStore(storage);
  assert.deepEqual(await store.list(), []);

  storage.setItem(storageKey, "{not json");
  assert.deepEqual(await store.list(), []);
});

test("malformed items are skipped without discarding valid favorites", async () => {
  const storage = new MemoryStorage();
  const valid = createFavorite({ futureGroup: "leadership" });
  storage.setItem(storageKey, JSON.stringify([valid, { id: 4 }, null]));
  const favorites = await new LocalStorageFavoritesStore(storage).list();
  assert.deepEqual(favorites, [valid]);
  assert.equal(favorites[0].futureGroup, "leadership");
});

test("service adds, deduplicates, edits, reorders, moves, and removes favorites", async () => {
  const storage = new MemoryStorage();
  const service = new FavoritesService(
    new LocalStorageFavoritesStore(storage),
    (() => {
      let id = 0;
      return () => `favorite-${++id}`;
    })(),
    () => new Date("2026-02-03T04:05:06.000Z"),
  );

  const first = await service.add("https://EXAMPLE.com/docs/steerco.pptx#slide=2", "Steerco", "Weekly");
  const duplicate = await service.add("https://example.com/docs/steerco.pptx", "Duplicate");
  assert.equal(duplicate.id, first.id);
  assert.equal((await service.list()).length, 1);

  const second = await service.add("https://example.com/docs/architecture.pptx");
  const third = await service.add("https://example.com/docs/template.pptx");
  await service.update(first.id, {
    name: "SUNRISE Steering Committee",
    comment: "Reference",
    url: first.url,
  });
  assert.deepEqual((await service.list())[0], {
    ...first,
    name: "SUNRISE Steering Committee",
    comment: "Reference",
    updatedAt: "2026-02-03T04:05:06.000Z",
  });

  await service.reorder([third.id, first.id, second.id]);
  assert.deepEqual((await service.list()).map(({ id }) => id), [third.id, first.id, second.id]);
  await service.move(first.id, -1);
  assert.deepEqual((await service.list()).map(({ id }) => id), [first.id, third.id, second.id]);
  await service.move(first.id, 1);
  assert.deepEqual((await service.list()).map(({ id }) => id), [third.id, first.id, second.id]);

  await service.remove(first.id);
  assert.deepEqual((await service.list()).map(({ id }) => id), [third.id, second.id]);
});

test("URL helpers normalize conservatively and derive decoded names", () => {
  assert.equal(
    normalizeFavoriteUrl(" HTTPS://Example.COM:443/docs/Quarterly%20Review.pptx/#slide=3 "),
    "https://example.com/docs/Quarterly%20Review.pptx",
  );
  assert.equal(
    deriveFavoriteName("https://example.com/docs/Quarterly%20Review.pptx?download=1"),
    "Quarterly Review",
  );
  assert.equal(
    deriveFavoriteName("https://example.com/docs/S4_Hosting_EA_Update_v1.0.pptx"),
    "S4 Hosting EA Update v1.0",
  );
  assert.equal(normalizeFavoriteUrl("not a URL "), "not a URL");
});

test("location helper produces short conservative hints", () => {
  assert.equal(deriveFavoriteLocation("https://danone-my.sharepoint.com/personal/user/Documents/deck.pptx"), "Personal OneDrive");
  assert.equal(deriveFavoriteLocation("https://danone.sharepoint.com/sites/WW-Cybersecurity/Shared/deck.pptx"), "WW-Cybersecurity");
  assert.equal(deriveFavoriteLocation("https://files.example.com/decks/file.pptx"), "files.example.com");
});

test("PowerPoint open URL launches the desktop client in edit mode", () => {
  assert.equal(
    getPowerPointOpenUrl("https://example.sharepoint.com/sites/team/Shared%20Documents/deck.pptx"),
    "ms-powerpoint:ofe|u|https://example.sharepoint.com/sites/team/Shared%20Documents/deck.pptx",
  );
  assert.throws(() => getPowerPointOpenUrl("file:///Users/example/deck.pptx"), /HTTP or HTTPS/);
});

test("filter searches name, comment, and URL without changing manual order", () => {
  const favorites = [
    createFavorite({ id: "1", name: "SUNRISE Steering Committee", comment: "Weekly" }),
    createFavorite({ id: "2", name: "Architecture Principles", comment: "Reference", url: "https://example.com/architecture.pptx" }),
    createFavorite({ id: "3", name: "GTI Review", url: "https://example.com/templates/gti-review.pptx" }),
  ];
  assert.deepEqual(filterFavorites(favorites, "reference").map(({ id }) => id), ["2"]);
  assert.deepEqual(filterFavorites(favorites, "STEERING").map(({ id }) => id), ["1"]);
  assert.deepEqual(filterFavorites(favorites, "templates").map(({ id }) => id), ["3"]);
  assert.deepEqual(filterFavorites(favorites, "").map(({ id }) => id), ["1", "2", "3"]);
});