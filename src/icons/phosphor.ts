export interface IconDefinition {
  name: string;
  slug: string;
  weight: string;
  path: string;
}

type IconIndex = Record<string, string[]>;

export interface PhosphorIconLibrary {
  icons: IconDefinition[];
  iconsPerWeight: number;
  weightCount: number;
}

function createIconDefinition(weight: string, filename: string): IconDefinition {
  const suffix = weight === "regular" ? "" : `-${weight}`;
  const slug = filename.slice(0, -4).replace(new RegExp(`${suffix}$`), "");

  return {
    name: slug.split("-").map((word) => word.charAt(0).toUpperCase() + word.slice(1)).join(" "),
    slug,
    weight,
    path: `assets/phosphor/phosphor-icons/SVGs/${weight}/${filename}`,
  };
}

export async function loadPhosphorIcons(): Promise<PhosphorIconLibrary> {
  const response = await fetch("assets/phosphor/icon-index.json");
  if (!response.ok) throw new Error("Could not load the Phosphor icon index.");

  const index = await response.json() as IconIndex;
  const icons = Object.entries(index).reduce<IconDefinition[]>(
    (allIcons, [weight, filenames]) => allIcons.concat(
      filenames.map((filename) => createIconDefinition(weight, filename)),
    ),
    [],
  );

  return {
    icons,
    iconsPerWeight: Object.values(index)[0]?.length ?? 0,
    weightCount: Object.keys(index).length,
  };
}

export async function insertPhosphorIcon(icon: IconDefinition, color: string): Promise<void> {
  const response = await fetch(icon.path);
  if (!response.ok) throw new Error(`Could not load ${icon.name}.`);

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
}