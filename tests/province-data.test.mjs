import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import test from "node:test";

const dataUrl = new URL("../src/data/provinces.json", import.meta.url);
const geoUrl = new URL("../public/data/provinces.geojson", import.meta.url);
const documentsUrl = new URL("../src/data/provincial-documents.json", import.meta.url);
const downloadsUrl = new URL("../public/downloads/planes-provinciales/", import.meta.url);

test("publishes one complete record for every province", async () => {
  const data = JSON.parse(await readFile(dataUrl, "utf8"));
  assert.equal(data.provinces.length, 32);
  assert.equal(new Set(data.provinces.map((item) => item.key)).size, 32);
  assert.equal(data.meta.totals.provincialPlans, 0);
  assert.equal(data.meta.totals.demands, 503);
  assert.equal(data.meta.totals.nationalInvestmentProjects, 2229);
  assert.ok(data.provinces.every((item) => item.plan.exists === false));
  assert.ok(data.provinces.every((item) => item.population > 0));
});

test("maps all province records to ADM1 geometry", async () => {
  const [data, geo] = await Promise.all([
    readFile(dataUrl, "utf8").then(JSON.parse),
    readFile(geoUrl, "utf8").then(JSON.parse),
  ]);
  assert.equal(geo.features.length, 32);
  const normalize = (value) =>
    String(value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "")
      .replace(/^bahoruco$/, "baoruco");
  const geometryNames = new Set(
    geo.features.map((feature) => normalize(feature.properties.shapeName)),
  );
  for (const province of data.provinces) {
    assert.ok(geometryNames.has(province.key), province.name);
  }
});

test("publishes one provincial-plan base document for all 32 provinces", async () => {
  const [provinceData, manifest, publishedFiles] = await Promise.all([
    readFile(dataUrl, "utf8").then(JSON.parse),
    readFile(documentsUrl, "utf8").then(JSON.parse),
    readdir(downloadsUrl),
  ]);
  assert.equal(manifest.documents.length, 32);
  assert.equal(new Set(manifest.documents.map((item) => item.provinceKey)).size, 32);
  assert.deepEqual(
    new Set(manifest.documents.map((item) => item.provinceKey)),
    new Set(provinceData.provinces.map((item) => item.key)),
  );
  assert.equal(publishedFiles.filter((name) => name.endsWith(".docx")).length, 32);

  for (const document of manifest.documents) {
    assert.match(document.territorialCode, /^\d{4}0000$/);
    assert.equal(document.territorialCode, `${document.regionCode}${document.provinceCode}0000`);
    assert.ok(document.fileName.startsWith(`${document.territorialCode}_Plan_Provincial_`));
    assert.ok(publishedFiles.includes(document.fileName));
    const bytes = await readFile(new URL(`../public/${document.path}`, import.meta.url));
    assert.ok(bytes.byteLength > 500_000, document.province);
    assert.equal(bytes.subarray(0, 2).toString("ascii"), "PK");
  }

  const appSource = await readFile(new URL("../src/App.jsx", import.meta.url), "utf8");
  assert.match(appSource, /Descargar documento base del Plan Provincial \(Word\)/);
  assert.match(appSource, /provincialDocuments\.get\(selected\.key\)/);
  assert.match(appSource, /¿Qué quiere comparar en el mapa\?/);
  assert.match(appSource, /Indicadores de contexto/);
  assert.match(appSource, /Instalaciones deportivas/);
  assert.doesNotMatch(appSource, /MTS_BASE_DOCUMENT/);
  assert.doesNotMatch(appSource, /Homicidios \/ 100 mil/);
  assert.doesNotMatch(appSource, /Planes Provinciales identificados: 0/);
  assert.doesNotMatch(appSource, /14000000_Plan_Provincial/);
  assert.doesNotMatch(appSource, /no constituye un plan aprobado/i);
});
