// Packs the built control (out/controls/SPFileExplorer) into a Dataverse solution zip without .NET,
// using the same layout as the SolutionPackager output of `dotnet build Solution/SPFileExplorerSolution`:
//
//   [Content_Types].xml, solution.xml, customizations.xml,
//   Controls/<prefix>_<namespace>.<constructor>/ControlManifest.xml, bundle.js, css/…, strings/…
//
// Usage:  npm run pack:solution            → Solution/SPFullFileExplorer.zip (managed)
//         npm run pack:solution -- --unmanaged → Solution/SPFullFileExplorer_unmanaged.zip
//
// The CI workflow builds the package with dotnet instead; both produce importable solutions.

import { readFileSync, readdirSync, statSync, writeFileSync, mkdirSync } from "node:fs";
import { join, relative, sep } from "node:path";
import { deflateRawSync } from "node:zlib";

const ROOT = new URL("..", import.meta.url).pathname;
const BUILD = join(ROOT, "out/controls/SPFileExplorer");
const SOLUTION_XML = join(ROOT, "Solution/SPFileExplorerSolution/src/Other/Solution.xml");
const MANIFEST_IN = join(ROOT, "SPFileExplorer/ControlManifest.Input.xml");
const unmanaged = process.argv.includes("--unmanaged");
const OUT = join(ROOT, unmanaged ? "Solution/SPFullFileExplorer_unmanaged.zip" : "Solution/SPFullFileExplorer.zip");
const BOM = "﻿";

// ── Versions and names ─────────────────────────────────────────────────────────────────────────────
const manifest = readFileSync(join(BUILD, "ControlManifest.xml"), "utf8");
const attr = (name) => {
  const m = new RegExp(`<control\\b[^>]*?\\s${name}="([^"]+)"`).exec(manifest);
  if (!m) throw new Error(`${name} not found in the built ControlManifest.xml. Run the production build first.`);
  return m[1];
};
const namespace = attr("namespace");
const constructorName = attr("constructor");
const controlVersion = attr("version");
const inputVersion = /<control\s[^>]*?\sversion="([^"]+)"/.exec(readFileSync(MANIFEST_IN, "utf8"))[1];
if (inputVersion !== controlVersion) {
  throw new Error(`Built control is ${controlVersion} but the manifest says ${inputVersion}. Rebuild first.`);
}

let solutionXml = readFileSync(SOLUTION_XML, "utf8").replace(/^﻿/, "");
const prefix = /<CustomizationPrefix>([^<]+)<\/CustomizationPrefix>/.exec(solutionXml)[1];
const controlName = `${prefix}_${namespace}.${constructorName}`;

// Same rule as .github/scripts/prepare-solution-version.mjs: <major>.<minor>.<ctrlMajor*100+ctrlMinor>.<ctrlPatch>
const [cMajor, cMinor, cPatch] = controlVersion.split(".").map(Number);
const sv = /<Version>(\d+)\.(\d+)[^<]*<\/Version>/.exec(solutionXml);
const solutionVersion = `${sv[1]}.${sv[2]}.${cMajor * 100 + cMinor}.${cPatch}`;
solutionXml = solutionXml
  .replace(sv[0], `<Version>${solutionVersion}</Version>`)
  .replace(/<Managed>\d<\/Managed>/, `<Managed>${unmanaged ? 0 : 1}</Managed>`)
  .replace(/<RootComponents\s*\/>|<RootComponents>[\s\S]*?<\/RootComponents>/,
    `<RootComponents>\n      <RootComponent type="66" schemaName="${controlName}" behavior="0" />\n    </RootComponents>`);

const customizationsXml = `<?xml version="1.0" encoding="utf-8"?>
<ImportExportXml xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <Entities />
  <Roles />
  <Workflows />
  <FieldSecurityProfiles />
  <Templates />
  <EntityMaps />
  <EntityRelationships />
  <OrganizationSettings />
  <optionsets />
  <CustomControls>
    <CustomControl>
      <Name>${controlName}</Name>
      <FileName>/Controls/${controlName}/ControlManifest.xml</FileName>
    </CustomControl>
  </CustomControls>
  <SolutionPluginAssemblies />
  <EntityDataProviders />
  <Languages>
    <Language>1033</Language>
  </Languages>
</ImportExportXml>`;

const contentTypes = `<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
  + `<Default Extension="xml" ContentType="text/xml" /><Default Extension="js" ContentType="application/octet-stream" />`
  + `<Default Extension="txt" ContentType="application/octet-stream" /><Default Extension="resx" ContentType="application/octet-stream" />`
  + `<Default Extension="css" ContentType="application/octet-stream" />`
  + `<Override PartName="/Controls/${controlName}/ControlManifest.xml" ContentType="application/octet-stream" /></Types>`;

// ── Files ──────────────────────────────────────────────────────────────────────────────────────────
const walk = (dir) => readdirSync(dir).flatMap((n) => {
  const p = join(dir, n);
  return statSync(p).isDirectory() ? walk(p) : [p];
});
const entries = [
  ["customizations.xml", Buffer.from(BOM + customizationsXml, "utf8")],
  ["solution.xml", Buffer.from(BOM + solutionXml, "utf8")],
  ...walk(BUILD).map((p) => [`Controls/${controlName}/${relative(BUILD, p).split(sep).join("/")}`, readFileSync(p)]),
  ["[Content_Types].xml", Buffer.from(BOM + contentTypes, "utf8")],
];

// ── Minimal zip writer (deflate) ─────────────────────────────────────────────────────────────────────
const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const now = new Date();
const dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xffff;
const dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xffff;

const locals = [];
const centrals = [];
let offset = 0;
for (const [name, data] of entries) {
  const nameBuf = Buffer.from(name, "utf8");
  const compressed = deflateRawSync(data, { level: 9 });
  const crc = crc32(data);
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6);
  local.writeUInt16LE(8, 8); local.writeUInt16LE(dosTime, 10); local.writeUInt16LE(dosDate, 12);
  local.writeUInt32LE(crc, 14); local.writeUInt32LE(compressed.length, 18); local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(nameBuf.length, 26); local.writeUInt16LE(0, 28);
  locals.push(local, nameBuf, compressed);
  const central = Buffer.alloc(46);
  central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
  central.writeUInt16LE(0x0800, 8); central.writeUInt16LE(8, 10); central.writeUInt16LE(dosTime, 12);
  central.writeUInt16LE(dosDate, 14); central.writeUInt32LE(crc, 16); central.writeUInt32LE(compressed.length, 20);
  central.writeUInt32LE(data.length, 24); central.writeUInt16LE(nameBuf.length, 28);
  central.writeUInt32LE(offset, 42);
  centrals.push(central, nameBuf);
  offset += local.length + nameBuf.length + compressed.length;
}
const centralSize = centrals.reduce((n, b) => n + b.length, 0);
const end = Buffer.alloc(22);
end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10);
end.writeUInt32LE(centralSize, 12); end.writeUInt32LE(offset, 16);

mkdirSync(join(ROOT, "Solution"), { recursive: true });
writeFileSync(OUT, Buffer.concat([...locals, ...centrals, end]));
console.log(`${relative(ROOT, OUT)}: ${controlName} ${controlVersion}, solution ${solutionVersion} (${unmanaged ? "unmanaged" : "managed"}), ${entries.length} files`);
