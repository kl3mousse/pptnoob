const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const source = path.join(root, "manifest.xml");
const outputDirectory = path.join(root, ".dev");
const output = path.join(outputDirectory, "manifest.xml");

const productionId = "df24bb71-4a1d-42e8-8382-07c4677bc8de";
const developmentId = "8e9a9f52-1d97-4cf4-82d7-3099777a0f91";
const productionOrigin = "https://kl3mousse.github.io/pptnoob";
const developmentOrigin = "https://localhost:3000";

const manifest = fs.readFileSync(source, "utf8")
  .replaceAll(productionId, developmentId)
  .replaceAll(productionOrigin, developmentOrigin)
  .replace('DefaultValue="pptNoob"', 'DefaultValue="pptNoob Dev"')
  .replace('id="Tab.Label" DefaultValue="pptNoob"', 'id="Tab.Label" DefaultValue="pptNoob Dev"');

if (!manifest.includes(developmentId) || !manifest.includes(developmentOrigin)) {
  throw new Error("Could not generate the development manifest from manifest.xml.");
}

fs.mkdirSync(outputDirectory, { recursive: true });
fs.writeFileSync(output, manifest);
console.log(`Generated ${path.relative(root, output)}`);