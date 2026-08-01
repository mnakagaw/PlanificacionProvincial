# Planificación Provincial

Portal cartográfico para organizar, por provincia, los insumos disponibles para
formular los primeros Planes Provinciales de la República Dominicana.

## Contenido

- mapa clicable de las 32 provincias y el Distrito Nacional;
- diagnóstico provincial resumido;
- proyectos y presupuesto asociado de Mapa Inversión 2026;
- demandas provinciales priorizadas;
- línea base de Plan Provincial: 0 planes publicados o identificados;
- enlaces a los cuatro portales DDPT utilizados como fuente.

SISMAP no forma parte del modelo provincial. La aplicación no interpreta los
indicadores municipales de SISMAP como avance de un Plan Provincial.

## Desarrollo

Requiere Node.js 22 o posterior.

```bash
npm install
npm run data:build
npm run dev
```

## Verificación

```bash
npm test
npm run build
```

La versión estática se genera en `dist/` con la ruta base:

`/DDPT/PlanificacionProvincial/`

## Publicación FTP

El despliegue reutiliza `FTP_HOST`, `FTP_USER`, `FTP_PASS` y
`FTP_REMOTE_ROOT` del archivo `.env` situado junto a este proyecto. El destino
está protegido para aceptar únicamente `DDPT/PlanificacionProvincial`.

```bash
npm run deploy:ftp
```

## Actualización de datos

`npm run data:build` reconstruye `src/data/provinces.json` y
`public/data/provinces.geojson` a partir de:

- Dashboard Territorial;
- Inversión Pública Territorial;
- Demandas Provinciales;
- cartografía municipal del Portal de Planificación Municipal, agrupada en ADM1.

Los conteos de demandas corresponden al consolidado 003 visible al 1 de agosto
de 2026. El script valida que existan exactamente 32 fichas y 32 geometrías.
