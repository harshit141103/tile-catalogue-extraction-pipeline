# Tile Catalogue Extraction Pipeline

## What this does

`run_pipeline.py` extracts tile records from a catalogue PDF into:

- `images/` - one clean PNG per detected tile image
- `output.xlsx` - exactly `Title | Length (mm) | Width (mm) | Image Name`
- `extraction_log.csv` - audit metadata for each record
- `validation_report.json` - validation checks and warnings

## Approach

The pipeline uses the PDF text layer for titles and dimensions, and the PDF image XObjects for tile artwork. It avoids full-page screenshots when a native embedded image can be decoded.

Workflow:

1. Read each page with `pdfplumber`.
2. Parse size strings such as `600x1200mm`, `600 × 600 mm`, and `600*1200`.
3. Find embedded raster images that look like product tiles.
4. Match each image to the first valid label line directly below it.
5. Assign dimensions by comparing the image aspect ratio to the page's extracted size metadata.
6. Save each tile as PNG with a deterministic filesystem-safe filename.
7. Upscale proportionally only when needed to satisfy the assignment's minimum pixel target.
8. Write the Excel file and validation report.

For wall-tile sections where the page header gives `300x600mm` or `300x450mm` but a companion tile image is square, the script records `300x300mm`. This is flagged in `extraction_log.csv` as a derived square companion size rather than hidden.

## How to run

From this folder:

```powershell
python run_pipeline.py "C:\Project\Nirwana\Copy of Odisha Catalogue .pdf" --output-dir .
```

Or from any folder:

```powershell
python "C:\Users\Harshit Jain\Documents\Codex\2026-08-23\operate-as-an-autonomous-coding-agent\outputs\submission\run_pipeline.py" "C:\Project\Nirwana\Copy of Odisha Catalogue .pdf" --output-dir "C:\Users\Harshit Jain\Documents\Codex\2026-08-23\operate-as-an-autonomous-coding-agent\outputs\submission"
```

Install dependencies if needed:

```powershell
python -m pip install -r requirements.txt
```

## Validation

The validation step checks:

- number of extracted records
- missing titles or dimensions
- duplicate title/dimension combinations
- duplicate or missing image files
- Excel row count versus image count
- exact Excel headers
- image dimensions
- aspect-ratio consistency against physical dimensions
- suspiciously small or blank images

## Assumptions and limitations

- The catalogue has selectable text for product labels and size metadata.
- Product titles are located directly below their tile image.
- Embedded image streams are the preferred image source.
- Scene mockups and decorative catalogue assets are skipped when they have no valid product label or only a footer page number.
- If a similar catalogue uses scanned pages without text, OCR would need to be added.
- If a page contains multiple dimension families, the assigned size is selected by aspect-ratio fit.
- Square companion pieces in wall-tile sets are assigned `300x300mm` from the stated 300 mm side plus square image aspect; this is logged per row.
