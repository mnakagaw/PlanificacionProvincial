import { useEffect, useMemo, useState } from "react";
import provinceData from "./data/provinces.json";

const { meta, provinces } = provinceData;

const REGION_ORDER = [
  "Cibao Noroeste",
  "Cibao Norte",
  "Cibao Nordeste",
  "Cibao Sur",
  "Valdesia",
  "El Valle",
  "Enriquillo",
  "Ozama",
  "Higuamo",
  "Yuma",
];

const REGION_COLORS = {
  "Cibao Noroeste": "#688b74",
  "Cibao Norte": "#3f7968",
  "Cibao Nordeste": "#70a19a",
  "Cibao Sur": "#91ad78",
  Valdesia: "#d3a24e",
  "El Valle": "#b68145",
  Enriquillo: "#cc765f",
  Ozama: "#355f78",
  Higuamo: "#5e87a0",
  Yuma: "#61a5a6",
};

const LAYERS = {
  overview: {
    label: "Panorama",
    eyebrow: "Territorio",
    description: "Las 10 regiones de planificación",
    color: "#1c5648",
  },
  population: {
    label: "Población",
    eyebrow: "Censo 2022",
    description: "Población provincial",
    color: "#337868",
  },
  investment: {
    label: "Inversión",
    eyebrow: "Mapa Inversión 2026",
    description: "Presupuesto asociado",
    color: "#b7792d",
  },
  demands: {
    label: "Demandas",
    eyebrow: "Agenda provincial",
    description: "Demandas priorizadas",
    color: "#3c7292",
  },
};

const PALETTES = {
  population: ["#e3efea", "#b8d5ca", "#84b6a4", "#4f907a", "#236650"],
  investment: ["#f8edd8", "#efd09a", "#dda95d", "#bd7b30", "#8a4f1e"],
  demands: ["#e7eff4", "#c1d6e3", "#8bb4cb", "#5b8fab", "#32677f"],
};

const numberFormatter = new Intl.NumberFormat("es-DO");
const compactFormatter = new Intl.NumberFormat("es-DO", {
  notation: "compact",
  maximumFractionDigits: 1,
});
const decimalFormatter = new Intl.NumberFormat("es-DO", {
  maximumFractionDigits: 1,
});

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

function projectPoint([longitude, latitude]) {
  const x = ((longitude + 72.05) / 3.95) * 1000;
  const y = ((19.95 - latitude) / 2.65) * 670;
  return [x, y];
}

function ringToPath(ring) {
  return ring
    .map((point, index) => {
      const [x, y] = projectPoint(point);
      return `${index === 0 ? "M" : "L"}${x.toFixed(2)},${y.toFixed(2)}`;
    })
    .join(" ")
    .concat(" Z");
}

function geometryToPath(geometry) {
  const polygons =
    geometry.type === "Polygon" ? [geometry.coordinates] : geometry.coordinates;
  return polygons.flatMap((polygon) => polygon.map(ringToPath)).join(" ");
}

function money(value) {
  if (!Number.isFinite(value)) return "No disponible";
  if (value >= 1_000_000_000) {
    return `RD$ ${decimalFormatter.format(value / 1_000_000_000)} mil M`;
  }
  return `RD$ ${decimalFormatter.format(value / 1_000_000)} M`;
}

function percent(value) {
  return Number.isFinite(value) ? `${decimalFormatter.format(value)}%` : "—";
}

function metricValue(province, layer) {
  if (layer === "population") return province.population;
  if (layer === "investment") return province.investment.budget;
  if (layer === "demands") return province.demands;
  return 0;
}

function metricLabel(province, layer) {
  if (layer === "population") return `${compactFormatter.format(province.population)} hab.`;
  if (layer === "investment") return money(province.investment.budget);
  if (layer === "demands") return `${numberFormatter.format(province.demands)} demandas`;
  return province.region;
}

function layerFill(province, layer, maximum) {
  if (layer === "overview") return REGION_COLORS[province.region] || "#93a39d";
  const palette = PALETTES[layer];
  if (!palette || maximum <= 0) return "#e2e7e4";
  const ratio = Math.sqrt(metricValue(province, layer) / maximum);
  const index = Math.min(
    palette.length - 1,
    Math.max(0, Math.round(ratio * (palette.length - 1))),
  );
  return palette[index];
}

function SourceLink({ href, label, meta: sourceMeta, primary = false }) {
  return (
    <a
      className={`source-link ${primary ? "is-primary" : ""}`}
      href={href}
      target="_blank"
      rel="noreferrer"
    >
      <span>
        <small>{sourceMeta}</small>
        <strong>{label}</strong>
      </span>
      <span aria-hidden="true">↗</span>
    </a>
  );
}

export function App() {
  const [activeLayer, setActiveLayer] = useState("overview");
  const [selectedRegion, setSelectedRegion] = useState("Todas");
  const [selected, setSelected] = useState(null);
  const [hovered, setHovered] = useState(null);
  const [mapShapes, setMapShapes] = useState([]);
  const [mapError, setMapError] = useState(false);

  const provinceLookup = useMemo(
    () => new Map(provinces.map((province) => [province.key, province])),
    [],
  );

  useEffect(() => {
    let current = true;
    fetch(`${import.meta.env.BASE_URL}data/provinces.geojson`)
      .then((response) => {
        if (!response.ok) throw new Error(`GeoJSON ${response.status}`);
        return response.json();
      })
      .then((collection) => {
        if (!current) return;
        setMapShapes(
          collection.features.map((feature) => ({
            key: canonicalProvince(feature.properties.shapeName),
            name: feature.properties.shapeName,
            path: geometryToPath(feature.geometry),
          })),
        );
      })
      .catch(() => current && setMapError(true));
    return () => {
      current = false;
    };
  }, []);

  const regions = useMemo(
    () => REGION_ORDER.filter((region) => provinces.some((item) => item.region === region)),
    [],
  );

  const filteredProvinces = useMemo(
    () =>
      provinces.filter(
        (province) => selectedRegion === "Todas" || province.region === selectedRegion,
      ),
    [selectedRegion],
  );

  const maximum = useMemo(
    () => Math.max(...provinces.map((province) => metricValue(province, activeLayer)), 0),
    [activeLayer],
  );

  const displayProvince = hovered || selected;
  const activeMeta = LAYERS[activeLayer];
  const totalPopulation = useMemo(
    () => provinces.reduce((total, province) => total + province.population, 0),
    [],
  );

  const layerTotals = {
    overview: numberFormatter.format(meta.totals.provinces),
    population: compactFormatter.format(totalPopulation),
    investment: numberFormatter.format(meta.totals.nationalInvestmentProjects),
    demands: numberFormatter.format(meta.totals.demands),
  };

  function chooseProvince(province) {
    setSelected(province);
    setSelectedRegion(province.region);
  }

  function clearSelection() {
    setSelected(null);
    setHovered(null);
    setSelectedRegion("Todas");
  }

  const investmentUrl = selected
    ? `${meta.sources.investment.replace("data/mapa_inversion.json", "")}?source=mapa&province=${encodeURIComponent(selected.name)}`
    : "https://prodecare.net/DDPT/InversionPublicaTerritorial/";

  return (
    <main className="portal">
      <header className="app-header">
        <a className="brand" href="#top" aria-label="Inicio">
          <span className="brand-mark" aria-hidden="true">PP</span>
          <span>
            <strong>Planificación Provincial</strong>
            <small>MHE · Dirección de Desarrollo y Planificación Territorial</small>
          </span>
        </a>
        <div className="header-meta">
          <span>Fuentes DDPT</span>
          <i aria-hidden="true" />
          <span>Corte {meta.investmentAsOf}</span>
        </div>
      </header>

      <section className="page-intro" id="top">
        <h1>Tablero de Planificación Provincial</h1>
        <p>
          Consulte el diagnóstico, la inversión pública, las demandas priorizadas
          y la situación del Plan Provincial.
        </p>
      </section>

      <section className="map-section">
        <div className="layer-switcher" role="group" aria-label="Lectura del mapa">
          {Object.entries(LAYERS).map(([key, layer]) => (
            <button
              type="button"
              key={key}
              className={activeLayer === key ? "is-active" : ""}
              onClick={() => setActiveLayer(key)}
              style={{ "--layer-accent": layer.color }}
              aria-pressed={activeLayer === key}
            >
              <i className="layer-dot" aria-hidden="true" />
              <span>{layer.label}</span>
              <strong>{layerTotals[key]}</strong>
            </button>
          ))}
          <article className="plan-summary" aria-label="Planes Provinciales identificados: 0">
            <i className="layer-dot" aria-hidden="true" />
            <span>Plan Provincial</span>
            <strong>{meta.totals.provincialPlans}</strong>
          </article>
        </div>

        <div className="selection-bar">
          <strong className="selection-title">Seleccione una provincia</strong>
          <label htmlFor="region-select">
            <span>Región</span>
            <select
              id="region-select"
              value={selectedRegion}
              onChange={(event) => {
                setSelectedRegion(event.target.value);
                setSelected(null);
              }}
            >
              <option value="Todas">Todas las regiones</option>
              {regions.map((region) => <option key={region}>{region}</option>)}
            </select>
          </label>
          <label htmlFor="province-select">
            <span>Provincia</span>
            <select
              id="province-select"
              value={selected?.key || ""}
              onChange={(event) => {
                const province = provinceLookup.get(event.target.value);
                if (province) chooseProvince(province);
                else setSelected(null);
              }}
            >
              <option value="">Seleccione una provincia</option>
              {filteredProvinces.map((province) => (
                <option key={province.key} value={province.key}>{province.name}</option>
              ))}
            </select>
          </label>
          <button type="button" className="clear-button" onClick={clearSelection}>
            <span aria-hidden="true">↺</span> Restablecer
          </button>
        </div>

        <div className="workspace">
          <article className="map-panel">
            <div className="map-header">
              <div>
                <i style={{ background: activeMeta.color }} aria-hidden="true" />
                <span>
                  <strong>{activeMeta.label}</strong>
                  <small>{activeMeta.description}</small>
                </span>
              </div>
              <p>Haga clic en el mapa para elegir una provincia</p>
            </div>
            <div className="map-canvas">
              {mapError ? (
                <div className="map-message">No fue posible cargar la cartografía.</div>
              ) : mapShapes.length === 0 ? (
                <div className="map-message">Cargando mapa provincial…</div>
              ) : (
                <svg
                  className="dominican-map"
                  viewBox="0 0 1000 670"
                  role="img"
                  aria-label="Mapa de las provincias de la República Dominicana"
                >
                  <defs>
                    <filter id="selected-province" x="-20%" y="-20%" width="140%" height="140%">
                      <feDropShadow dx="0" dy="4" stdDeviation="5" floodColor="#0d2922" floodOpacity="0.32" />
                    </filter>
                  </defs>
                  {mapShapes.map((shape) => {
                    const province = provinceLookup.get(shape.key);
                    if (!province) return null;
                    const inScope = selectedRegion === "Todas" || province.region === selectedRegion;
                    const isSelected = selected?.key === province.key;
                    return (
                      <path
                        key={shape.key}
                        d={shape.path}
                        fill={inScope ? layerFill(province, activeLayer, maximum) : "#dfe4e1"}
                        fillRule="evenodd"
                        stroke={isSelected ? "#f4c55a" : "#ffffff"}
                        strokeWidth={isSelected ? 3.2 : 1.15}
                        vectorEffect="non-scaling-stroke"
                        filter={isSelected ? "url(#selected-province)" : undefined}
                        className={`province-shape ${inScope ? "" : "is-muted"}`}
                        tabIndex="0"
                        onMouseEnter={() => setHovered(province)}
                        onMouseLeave={() => setHovered(null)}
                        onFocus={() => setHovered(province)}
                        onBlur={() => setHovered(null)}
                        onClick={() => chooseProvince(province)}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === " ") chooseProvince(province);
                        }}
                        aria-label={`${province.name}: ${metricLabel(province, activeLayer)}`}
                      >
                        <title>{province.name} · {metricLabel(province, activeLayer)}</title>
                      </path>
                    );
                  })}
                </svg>
              )}

              {displayProvince && (
                <div className="map-tooltip" aria-live="polite">
                  <small>{displayProvince.region}</small>
                  <strong>{displayProvince.name}</strong>
                  <span>{metricLabel(displayProvince, activeLayer)}</span>
                </div>
              )}

              <div className="map-legend">
                {activeLayer === "overview" ? (
                  regions.map((region) => (
                    <span key={region}><i style={{ background: REGION_COLORS[region] }} />{region}</span>
                  ))
                ) : (
                  <>
                    <span>Bajo</span>
                    <span className="gradient" style={{ "--gradient": `linear-gradient(90deg, ${PALETTES[activeLayer].join(",")})` }} />
                    <span>Alto</span>
                  </>
                )}
              </div>
            </div>
            <div className="map-footer">
              <span>Cartografía: división político-administrativa ADM1</span>
              <span>Seleccione una provincia para ver sus datos</span>
            </div>
          </article>

          <aside className="detail-panel">
            {selected ? (
              <>
                <div className="detail-heading">
                  <span>{selected.region}</span>
                  <h2>{selected.name}</h2>
                  <p>{numberFormatter.format(selected.population)} habitantes · Censo 2022</p>
                </div>

                <div className="plan-status">
                  <span aria-hidden="true">PP</span>
                  <div>
                    <small>PLAN PROVINCIAL</small>
                    <strong>{selected.plan.status}</strong>
                    <p>No se ha identificado un plan provincial publicado.</p>
                  </div>
                </div>

                <section className="detail-section">
                  <div className="detail-section-title">
                    <span>Diagnóstico</span>
                    <small>Indicadores disponibles</small>
                  </div>
                  <div className="metric-grid">
                    <div><span>Municipios</span><strong>{selected.municipalityCount}</strong></div>
                    <div><span>Homicidios / 100 mil</span><strong>{decimalFormatter.format(selected.diagnosis.homicideRate)}</strong><small>{selected.diagnosis.homicideYear}</small></div>
                    <div><span>Hacinamiento extremo</span><strong>{percent(selected.diagnosis.extremeOvercrowdingPct)}</strong><small>{selected.diagnosis.overcrowdingYear}</small></div>
                    <div><span>Centros INAIPI</span><strong>{numberFormatter.format(selected.diagnosis.inaipiCenters)}</strong></div>
                  </div>
                </section>

                <section className="detail-section investment-summary">
                  <div className="detail-section-title">
                    <span>Inversión pública</span>
                    <small>{selected.investment.year}</small>
                  </div>
                  <strong className="money-value">{money(selected.investment.budget)}</strong>
                  <p>{selected.investment.projectCount} proyectos asociados · {selected.investment.topSector} es el sector más frecuente.</p>
                  <div className="progress-row">
                    <span><small>Ejecución registrada</small><strong>{percent(selected.investment.executionPct)}</strong></span>
                    <i><b style={{ width: `${Math.min(selected.investment.executionPct, 100)}%` }} /></i>
                  </div>
                </section>

                <section className="demand-summary">
                  <span>Demandas priorizadas</span>
                  <strong>{selected.demands}</strong>
                  <p>{selected.demands === 0 ? "Sin registros en el consolidado vigente." : "registros en la agenda provincial consolidada."}</p>
                </section>

                <div className="source-links">
                  <SourceLink href="https://prodecare.net/DDPT/Dashboard-Territorial/" label="Abrir diagnóstico territorial" meta="Población, servicios y contexto" primary />
                  <SourceLink href={investmentUrl} label="Ver inversión de la provincia" meta="Mapa Inversión 2026" />
                  <SourceLink href={meta.sources.demands} label="Consultar demandas provinciales" meta="Consolidado por institución" />
                  <SourceLink href={meta.sources.municipalPlanning} label="Revisar planes municipales" meta="Insumos de los municipios" />
                </div>
              </>
            ) : (
              <div className="empty-detail">
                <span className="empty-marker" aria-hidden="true">⌖</span>
                <span className="empty-kicker">FICHA PROVINCIAL</span>
                <h2>Seleccione una provincia</h2>
                <p>
                  Haga clic en el mapa o utilice los selectores para consultar
                  diagnóstico, inversión, demandas y estado del Plan Provincial.
                </p>
                <div className="empty-list">
                  <span>01 <b>Diagnóstico territorial</b></span>
                  <span>02 <b>Inversión pública 2026</b></span>
                  <span>03 <b>Demandas priorizadas</b></span>
                </div>
              </div>
            )}
          </aside>
        </div>
      </section>

      <section className="method-section">
        <div className="section-heading light">
          <div>
            <span className="section-index">02 / RUTA DE FORMULACIÓN</span>
            <h2>De los datos al Plan Provincial</h2>
          </div>
          <p>Los tres portales existentes funcionan como insumos, no como sustituto del proceso participativo provincial.</p>
        </div>
        <div className="method-grid">
          <article><span>01</span><small>LEER</small><h3>Diagnóstico</h3><p>Caracterizar población, servicios, economía y brechas territoriales.</p></article>
          <article><span>02</span><small>CONTRASTAR</small><h3>Inversión</h3><p>Revisar proyectos, presupuesto y ejecución pública en el territorio.</p></article>
          <article><span>03</span><small>PRIORIZAR</small><h3>Demandas</h3><p>Ordenar las demandas por tema, institución y alcance provincial.</p></article>
          <article className="method-final"><span>04</span><small>FORMULAR</small><h3>Plan Provincial</h3><p>Construir visión, objetivos, cartera y mecanismos de seguimiento.</p></article>
        </div>
      </section>

      <footer>
        <div className="footer-brand"><span>PP</span><strong>Planificación Provincial</strong></div>
        <p>
          Fuentes integradas: Dashboard Territorial, Inversión Pública Territorial,
          Demandas Provinciales y Planificación Municipal. Los montos asociados a
          proyectos multiterritoriales pueden representar el presupuesto completo del proyecto.
        </p>
        <span>Corte de inversión: {meta.investmentAsOf}</span>
      </footer>
    </main>
  );
}
