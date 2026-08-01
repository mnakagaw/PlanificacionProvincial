import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import polygonClipping from "polygon-clipping";

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");

const SOURCES = {
  diagnostic:
    "https://prodecare.net/DDPT/Dashboard-Territorial/data/territorial-dashboard.json",
  investment:
    "https://prodecare.net/DDPT/InversionPublicaTerritorial/data/mapa_inversion.json",
  demands:
    "https://prodecare.net/DDPT/DemandasProvinciales/",
  municipalPlanning:
    "https://prodecare.net/DDPT/PlanificacionMunicipal/",
  geometry:
    "https://prodecare.net/DDPT/PlanificacionMunicipal/data/adm2.geojson",
};

const demandCounts = {
  "Monte Cristi": 14,
  Dajabón: 15,
  "Elías Piña": 18,
  Independencia: 16,
  Pedernales: 22,
  "La Altagracia": 16,
  "El Seibo": 15,
  "Hato Mayor": 18,
  Samaná: 15,
  "María Trinidad Sánchez": 10,
  Espaillat: 10,
  "Puerto Plata": 25,
  Barahona: 12,
  Azua: 14,
  Peravia: 12,
  "San Cristóbal": 18,
  "San José de Ocoa": 9,
  "Santo Domingo": 0,
  "Distrito Nacional": 1,
  "San Pedro de Macorís": 21,
  "La Romana": 12,
  Valverde: 21,
  "Santiago Rodríguez": 13,
  "La Vega": 24,
  "Hermanas Mirabal": 9,
  Duarte: 19,
  Santiago: 21,
  Baoruco: 16,
  "San Juan": 23,
  "Monseñor Nouel": 37,
  "Sánchez Ramírez": 10,
  "Monte Plata": 17,
};

function normalize(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("es")
    .replace(/[^a-z0-9]/g, "");
}

function canonicalProvince(value) {
  const cleaned = normalize(value);
  return cleaned === "bahoruco" ? "baoruco" : cleaned;
}

async function fetchJson(url) {
  const response = await fetch(url, { redirect: "follow" });
  if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`);
  return response.json();
}

const [diagnostic, investment, municipalGeometry] = await Promise.all([
  fetchJson(SOURCES.diagnostic),
  fetchJson(SOURCES.investment),
  fetchJson(SOURCES.geometry),
]);

const groupedGeometry = new Map();
for (const feature of municipalGeometry.features) {
  const name = feature.properties?.provincia;
  if (!name) continue;
  const key = canonicalProvince(name);
  if (!groupedGeometry.has(key)) {
    groupedGeometry.set(key, { name, geometries: [] });
  }
  const target = groupedGeometry.get(key);
  target.geometries.push(feature.geometry.coordinates);
}

const geometry = {
  type: "FeatureCollection",
  features: [...groupedGeometry.entries()].map(([key, item]) => ({
    type: "Feature",
    properties: {
      shapeName: item.name,
      provinceKey: key,
      shapeType: "ADM1",
    },
    geometry: {
      type: "MultiPolygon",
      coordinates: polygonClipping.union(...item.geometries),
    },
  })),
};

const demandLookup = new Map(
  Object.entries(demandCounts).map(([name, count]) => [
    canonicalProvince(name),
    count,
  ]),
);

const provinces = diagnostic.provinces
  .map((province) => {
    const key = canonicalProvince(province.name);
    const projects = investment.projects.filter((project) =>
      (project.provinces || []).some(
        (name) => canonicalProvince(name) === key,
      ),
    );
    const investmentBudget = projects.reduce(
      (total, project) => total + (Number(project.budget) || 0),
      0,
    );
    const investmentExecuted = projects.reduce(
      (total, project) => total + (Number(project.executed) || 0),
      0,
    );
    const sectorCounts = new Map();
    for (const project of projects) {
      for (const sector of project.sectors || []) {
        sectorCounts.set(sector, (sectorCounts.get(sector) || 0) + 1);
      }
    }
    const topSector = [...sectorCounts.entries()].sort(
      (a, b) => b[1] - a[1],
    )[0]?.[0];

    return {
      key,
      name: province.name,
      region:
        province.region === "Ozama o Metropolitana"
          ? "Ozama"
          : province.region,
      population: province.population,
      municipalityCount: province.municipalityCount,
      plan: {
        exists: false,
        status: "Por formular",
      },
      diagnosis: {
        homicideRate: province.metrics?.homicide?.latest?.rate ?? null,
        homicideYear: province.metrics?.homicide?.latest?.year ?? null,
        extremeOvercrowdingPct:
          province.metrics?.overcrowding?.extremePct ?? null,
        overcrowdingYear: province.metrics?.overcrowding?.year ?? null,
        inaipiCenters: province.metrics?.inaipi?.centers ?? null,
        sportsFacilities: province.metrics?.sports?.count ?? null,
      },
      investment: {
        year: investment.meta.year,
        projectCount: projects.length,
        budget: investmentBudget,
        executed: investmentExecuted,
        executionPct:
          investmentBudget > 0
            ? (investmentExecuted / investmentBudget) * 100
            : 0,
        projectsWithExecution: projects.filter(
          (project) => Number(project.executed) > 0,
        ).length,
        projectsWithActiveContracts: projects.filter(
          (project) => Number(project.activeContracts) > 0,
        ).length,
        topSector: topSector || "Sin sector identificado",
      },
      demands: demandLookup.get(key) ?? 0,
    };
  })
  .sort((a, b) => a.name.localeCompare(b.name, "es"));

const missingGeometry = provinces.filter(
  (province) =>
    !geometry.features.some(
      (feature) =>
        canonicalProvince(feature.properties?.shapeName) === province.key,
    ),
);

if (provinces.length !== 32) {
  throw new Error(`Se esperaban 32 provincias y llegaron ${provinces.length}.`);
}
if (missingGeometry.length) {
  throw new Error(
    `Falta cartografía: ${missingGeometry.map((item) => item.name).join(", ")}`,
  );
}

const output = {
  meta: {
    generatedAt: new Date().toISOString(),
    diagnosticGeneratedAt: diagnostic.meta.generated,
    investmentAsOf: investment.meta.asOf,
    investmentYear: investment.meta.year,
    planBaseline: "2026-08-01",
    sources: SOURCES,
    totals: {
      provinces: provinces.length,
      provincialPlans: 0,
      demands: provinces.reduce((total, item) => total + item.demands, 0),
      nationalInvestmentProjects: investment.meta.totals.projects,
    },
  },
  provinces,
};

await mkdir(resolve(projectRoot, "src/data"), { recursive: true });
await mkdir(resolve(projectRoot, "public/data"), { recursive: true });
await writeFile(
  resolve(projectRoot, "src/data/provinces.json"),
  `${JSON.stringify(output, null, 2)}\n`,
  "utf8",
);
await writeFile(
  resolve(projectRoot, "public/data/provinces.geojson"),
  `${JSON.stringify(geometry)}\n`,
  "utf8",
);

console.log(
  `Datos creados: ${provinces.length} provincias, ${output.meta.totals.demands} demandas.`,
);
