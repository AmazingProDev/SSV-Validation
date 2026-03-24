# SSV Drive Test Analyzer

A Flask web app for analyzing **2G/3G/LTE/5G drive test reports** (SSV Excel files). Automatically detects cross patterns in serving cell maps and degradation zones in coverage/quality/throughput maps.

## Features

- **Multi-sheet Excel parsing** — automatically extracts images from all sheets
- **Image classification** — identifies Serving Cell (PCI), Coverage (RSRP/RSCP), Quality (SINR/RSRQ), and throughput maps
- **Cross detection** — analyzes serving cell maps for 3-sector overlap issues
  - Returns **CROSS DETECTED** (red) or **NO CROSS DETECTED** (green)
- **Degradation zone detection** — identifies red dot clusters in coverage/quality/throughput maps
  - Draws magenta circles around degraded areas
  - Supports sparse (6+ dots) to dense degradation patterns
- **Tabbed results UI** — organized by sheet with summary page showing overall SSV status

## Deployment

### Local Development

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

### Vercel Deployment

The app is **Vercel-ready**. Images are embedded as base64 in responses (no persistent `/tmp` files).

```bash
vercel --prod
```

**Requirements:**
- Vercel **Pro tier** recommended (10s timeout on Hobby may be insufficient for large files)
- Python 3.9+

## Files

- **`app.py`** — Flask backend with all analysis logic
- **`templates/index.html`** — Dark-themed UI with tabbed results
- **`requirements.txt`** — Python dependencies
- **`vercel.json`** — Vercel configuration
- **`.gitignore`** / **`.vercelignore`** — excludes test data and cache

## Technical Details

### Cross Detection Algorithm
1. Estimate site center from high-density colorful pixels
2. Detect 3 dominant hue peaks in a radial fan around the site
3. Assign each colored pixel to a sector by color similarity
4. Apply connected-component filtering (keep only 5–90px blobs, >0.2 fill ratio)
5. Measure misassignment, mixed-bin, and intrusion ratios
6. Return **CROSS DETECTED** if any ratio exceeds threshold

### Red Zone Detection
- Strict HSV thresholds: H∈[0,8]∪[172,180], S>150, V>150 (pure red only)
- Morphological opening (5×5 ellipse) to remove thin text
- Union-Find clustering (max distance 8% of image diagonal)
- Filter circles by size (~100–4000px²) and fill ratio

## Known Limitations

- **Vercel timeout:** Pro tier (300s) recommended for >100 images per file
- **Image embedding:** responses limited to ~4.5MB (OK for typical SSV files)
- **Language:** keywords in French/English; extend `IMAGE_TYPE_KEYWORDS` for other languages

## License

MIT
