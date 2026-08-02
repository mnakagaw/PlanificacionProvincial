import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const dataUrl = new URL("../src/data/provinces.json", import.meta.url);
const geoUrl = new URL("../public/data/provinces.geojson", import.meta.url);

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

test("publishes the MTS provincial-plan base document without municipal-only references", async () => {
  const documentUrl = new URL(
    "../public/downloads/planes-provinciales/03140000_Plan_Provincial_Maria_Trinidad_Sanchez_Documento_Base_2026.docx",
    import.meta.url,
  );
  const bytes = await readFile(documentUrl);
  assert.ok(bytes.byteLength > 500_000);
  assert.equal(bytes.subarray(0, 2).toString("ascii"), "PK");

  const appSource = await readFile(new URL("../src/App.jsx", import.meta.url), "utf8");
  assert.match(appSource, /Descargar documento base del Plan Provincial \(Word\)/);
  assert.match(appSource, /mariatrinidadsanchez/);
  assert.match(appSource, /03140000_Plan_Provincial/);
  assert.doesNotMatch(appSource, /14000000_Plan_Provincial/);
  assert.doesNotMatch(appSource, /no constituye un plan aprobado/i);
});
