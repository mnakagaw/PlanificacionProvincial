from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from PIL import Image


PAGE_RE = re.compile(r"page-(\d+)\.png$", re.IGNORECASE)
EXPECTED_SIZES = {(1020, 1320), (1320, 1020)}


def main():
    parser = argparse.ArgumentParser(description="Valida todas las páginas PNG renderizadas.")
    parser.add_argument("directory", type=Path)
    args = parser.parse_args()
    root = args.directory.resolve()

    documents = 0
    pages_total = 0
    portrait = 0
    landscape = 0
    anomalies = []
    page_counts = {}

    for document_dir in sorted(path for path in root.iterdir() if path.is_dir()):
        pages = sorted(
            (path for path in document_dir.glob("page-*.png") if PAGE_RE.search(path.name)),
            key=lambda path: int(PAGE_RE.search(path.name).group(1)),
        )
        if not pages:
            anomalies.append({"document": document_dir.name, "issue": "sin páginas"})
            continue
        documents += 1
        page_counts[document_dir.name] = len(pages)
        expected_numbers = list(range(1, len(pages) + 1))
        actual_numbers = [int(PAGE_RE.search(path.name).group(1)) for path in pages]
        if actual_numbers != expected_numbers:
            anomalies.append({"document": document_dir.name, "issue": "secuencia incompleta"})

        for page in pages:
            pages_total += 1
            with Image.open(page).convert("L") as image:
                if image.size not in EXPECTED_SIZES:
                    anomalies.append({"page": str(page), "issue": f"tamaño {image.size}"})
                elif image.width < image.height:
                    portrait += 1
                else:
                    landscape += 1

                ink = image.point(lambda value: 255 if value < 248 else 0)
                bbox = ink.getbbox()
                if bbox is None:
                    anomalies.append({"page": str(page), "issue": "página en blanco"})
                    continue
                if bbox[0] <= 2 or bbox[1] <= 2 or bbox[2] >= image.width - 2 or bbox[3] >= image.height - 2:
                    anomalies.append({"page": str(page), "issue": f"contenido toca borde {bbox}"})
                ink_pixels = ink.histogram()[255]
                if ink_pixels / (image.width * image.height) < 0.001:
                    anomalies.append({"page": str(page), "issue": "contenido excepcionalmente escaso"})

    print(
        json.dumps(
            {
                "documents": documents,
                "pages": pages_total,
                "portrait_pages": portrait,
                "landscape_pages": landscape,
                "minimum_pages": min(page_counts.values()) if page_counts else 0,
                "maximum_pages": max(page_counts.values()) if page_counts else 0,
                "anomalies": anomalies,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    if anomalies:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
