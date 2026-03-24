import math
import os
import shutil
import zipfile
import uuid
import traceback
import base64
import cv2
import numpy as np
import xml.etree.ElementTree as ET
from collections import defaultdict
from flask import Flask, render_template, request, send_from_directory, jsonify, Response
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Vercel's filesystem is read-only except /tmp.
# VERCEL env var is set automatically by the Vercel runtime.
_UPLOAD_ROOT = '/tmp' if os.environ.get('VERCEL') else os.path.dirname(os.path.abspath(__file__))
app.config['UPLOAD_FOLDER'] = os.path.join(_UPLOAD_ROOT, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 MB

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# Cross-check algorithm constants  (mirrored from analyzer.py)
# ─────────────────────────────────────────────────────────────────────────────
_CC_SAT_THR    = 0.50   # colorful mask – min saturation [0,1]
_CC_VAL_THR    = 0.35   # colorful mask – min value [0,1]
_SITE_HIST_XR  = (0.45, 0.75)
_SITE_HIST_YR  = (0.10, 0.55)
_SITE_PRIOR_XR = 0.60
_SITE_PRIOR_YR = 0.32
_SITE_WIN_R    = 16     # density window radius for site center estimation
_SITE_MIN_DEN  = 35     # min density to be a site center candidate
_SITE_N_SAMP   = 90     # max candidates to average
_SITE_DW       = 2.5    # distance-to-prior penalty weight
_SITE_FLOOR    = 0.82   # score floor ratio
_FAN_INNER     = 6      # inner radius of sector-hue fan
_FAN_OUTER     = 36     # outer radius of sector-hue fan
_FAN_TARGET    = 18     # target radius (optimal weight)
_FAN_DEN_R     = 5      # density radius inside the fan
_FAN_MIN_DEN   = 22     # min density inside the fan
_HUE_PEAK_MIN  = 180    # min weighted count for a hue peak (30 × 6)
_DENSE_WIN_FB  = 7      # fallback: density window for hue search
_HUE_PEAK_FB   = 30     # fallback: min count per bin
_PNT_WIN_R     = 3      # density window radius for point candidates
_PNT_MIN_D     = 13     # min density for a point
_PNT_MAX_D     = 45     # max density for a point (excludes solid blobs)
_PNT_MIN_R     = 22     # min distance from site to count as a point
_MIN_PER_CLR   = 80     # min point pixels per sector
_MIN_TOTAL     = 300    # min total point pixels across all sectors
_PROTO_HUE_WIN = 10.0   # hue window (deg) when sampling sector prototypes
_HUE_CAP_DEG   = 18.0
_HUE_FLOOR_DEG = 8.0
_RGB_CAP       = 80.0
_RGB_FLOOR     = 45.0
_MISASSIGN_THR = 0.18
_MIXED_THR     = 0.20
_INTRUSION_THR = 0.18
_MISASSIGN_MAR = 12.0   # degree margin for misassignment check

# ─────────────────────────────────────────────────────────────────────────────
# Namespaces
# ─────────────────────────────────────────────────────────────────────────────
MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XDR_NS  = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A_NS    = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# ─────────────────────────────────────────────────────────────────────────────
# Image type classification
# ─────────────────────────────────────────────────────────────────────────────

# Order matters: more specific keywords first
IMAGE_TYPE_KEYWORDS = [
    ('dl_throughput', ['débit dl', 'dl throughput', 'downlink', 'rlc dl', 'download', 'débit dl en']),
    ('ul_throughput', ['débit ul', 'ul throughput', 'uplink', 'rlc ul', 'upload', 'débit ul en']),
    ('serving_cell',  ['pci de', 'pci ', 'serving cell', 'cellule serveuse', 'sc pci',
                       'cell serveuse', 'cell id', 'sc freq', 'pilote',
                       'sweving pci', 'serving pci', 'best server',
                       'bcch', 'bsic', ' lac ', 'lac id',      # 2G GSM (space-bounded to avoid matching "emplacement")
                       'cells pci', 'cellules serveuse']),      # variants
    ('coverage',      ['rsrp', 'rscp', 'rxlev', 'couverture', 'coverage', 'signal level']),
    ('quality',       ['sinr', 'rsrq', 'rxqual', 'qualité', 'quality', 'snr', 'ec/io', 'ec/n0']),
    ('geo',           ['emplacement', 'géographique', 'geographic', 'location', 'géo']),
]

IMAGE_TYPE_META = {
    'geo':           {'label': 'Geographic Location', 'icon': '🗺',  'color': '#64748b'},
    'serving_cell':  {'label': 'Serving Cell (PCI)',  'icon': '📡', 'color': '#8b5cf6'},
    'coverage':      {'label': 'Coverage',            'icon': '📶', 'color': '#0ea5e9'},
    'quality':       {'label': 'Quality',             'icon': '⭐', 'color': '#f59e0b'},
    'dl_throughput': {'label': 'DL Throughput',       'icon': '⬇',  'color': '#10b981'},
    'ul_throughput': {'label': 'UL Throughput',       'icon': '⬆',  'color': '#06b6d4'},
    'unknown':       {'label': 'Drive Test Map',      'icon': '📍', 'color': '#94a3b8'},
}

# Sheets to skip (logos, raw data, screenshots)
SKIP_SHEET_KEYWORDS = ['cover', 'screen', 'données enodeb', 'enodeb', 'tests statiques']


def should_skip_sheet(name):
    nl = name.lower().strip()
    return any(kw in nl for kw in SKIP_SHEET_KEYWORDS)


def col_letter_to_num(col_str):
    """Excel column letter → 1-based integer (A=1, Z=26, AA=27…)."""
    n = 0
    for c in col_str.upper():
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n


def classify_image_type(label):
    ll = label.lower()
    for img_type, keywords in IMAGE_TYPE_KEYWORDS:
        if any(kw in ll for kw in keywords):
            return img_type
    return 'unknown'


# ─────────────────────────────────────────────────────────────────────────────
# Excel parsing — extract images with labels
# ─────────────────────────────────────────────────────────────────────────────

def parse_excel_sheets(xlsx_path, output_dir):
    """
    Parse the xlsx file.
    For each DT sheet, read the drawing XML to find image placements,
    resolve nearby cell labels to classify each image, and extract it.

    Returns:
        [{'sheet_name': str, 'images': [ImageInfo]}]
    where ImageInfo = {filename, path, label, type, from_row, from_col}
    """
    results = []

    with zipfile.ZipFile(xlsx_path, 'r') as z:
        all_files = set(z.namelist())

        # ── workbook: sheet list ──────────────────────────────────────────
        wb_root = ET.fromstring(z.read('xl/workbook.xml'))

        wb_rels = {}
        for rel in ET.fromstring(z.read('xl/_rels/workbook.xml.rels')):
            wb_rels[rel.get('Id')] = rel.get('Target')

        # ── shared strings ────────────────────────────────────────────────
        shared_strings = []
        if 'xl/sharedStrings.xml' in all_files:
            ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in ss_root.findall(f'{{{MAIN_NS}}}si'):
                text = ''.join(t.text or '' for t in si.findall(f'.//{{{MAIN_NS}}}t'))
                shared_strings.append(text)

        # ── process each sheet ────────────────────────────────────────────
        for sheet_elem in wb_root.findall(f'.//{{{MAIN_NS}}}sheet'):
            sheet_name = sheet_elem.get('name', '')
            if should_skip_sheet(sheet_name):
                continue

            sheet_rid  = sheet_elem.get(f'{{{REL_NS}}}id')
            sheet_file = wb_rels.get(sheet_rid, '')
            sheet_path = f'xl/{sheet_file}'
            if sheet_path not in all_files:
                continue

            # Sheet relationships
            sheet_rels_path = (
                sheet_path
                .replace('worksheets/', 'worksheets/_rels/')
                .replace('.xml', '.xml.rels')
            )
            if sheet_rels_path not in all_files:
                continue

            drawing_target = None
            for rel in ET.fromstring(z.read(sheet_rels_path)):
                if rel.get('Type', '').split('/')[-1] == 'drawing':
                    drawing_target = rel.get('Target', '')
            if not drawing_target:
                continue

            # Drawing file path (relative to xl/worksheets/)
            drawing_file = 'xl/' + drawing_target.lstrip('../')
            if drawing_file not in all_files:
                continue

            drawing_rels_file = (
                drawing_file
                .replace('drawings/', 'drawings/_rels/')
                .replace('.xml', '.xml.rels')
            )
            if drawing_rels_file not in all_files:
                continue

            # drawing rId → image filename
            dr_rels = {}
            for rel in ET.fromstring(z.read(drawing_rels_file)):
                if rel.get('Type', '').split('/')[-1] == 'image':
                    target = rel.get('Target', '').split('/')[-1]
                    dr_rels[rel.get('Id')] = target

            # Cell text map: (row_1based, col_1based) → string
            sheet_cells = {}
            for row_el in ET.fromstring(z.read(sheet_path)).findall(f'.//{{{MAIN_NS}}}row'):
                row_num = int(row_el.get('r', 0))
                for cell_el in row_el.findall(f'{{{MAIN_NS}}}c'):
                    if cell_el.get('t') == 's':
                        v_el = cell_el.find(f'{{{MAIN_NS}}}v')
                        if v_el is not None:
                            try:
                                idx = int(v_el.text)
                                if 0 <= idx < len(shared_strings):
                                    cell_ref = cell_el.get('r', '')
                                    col_letters = ''.join(c for c in cell_ref if c.isalpha())
                                    if col_letters:
                                        sheet_cells[(row_num, col_letter_to_num(col_letters))] = shared_strings[idx]
                            except (ValueError, TypeError):
                                pass

            # Parse drawing anchors
            drawing_root = ET.fromstring(z.read(drawing_file))
            sheet_images = []

            for anchor in drawing_root.findall(f'{{{XDR_NS}}}twoCellAnchor'):
                from_col_el = anchor.find(f'{{{XDR_NS}}}from/{{{XDR_NS}}}col')
                from_row_el = anchor.find(f'{{{XDR_NS}}}from/{{{XDR_NS}}}row')
                if from_col_el is None or from_row_el is None:
                    continue

                from_col = int(from_col_el.text)  # 0-based
                from_row = int(from_row_el.text)   # 0-based

                blip = anchor.find(f'.//{{{A_NS}}}blip')
                if blip is None:
                    continue
                rid = blip.get(f'{{{REL_NS}}}embed')
                if rid not in dr_rels:
                    continue

                img_filename = dr_rels[rid]
                img_src = f'xl/media/{img_filename}'
                if img_src not in all_files:
                    continue

                img_data = z.read(img_src)
                img_dest = os.path.join(output_dir, img_filename)
                with open(img_dest, 'wb') as f:
                    f.write(img_data)

                # Skip logos / header images (< 200×100 px)
                nparr = np.frombuffer(img_data, np.uint8)
                img_cv = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                if img_cv is None:
                    continue
                img_h, img_w = img_cv.shape[:2]
                if img_h < 100 or img_w < 200:
                    continue

                # Find label in nearby cells
                excel_row = from_row + 1  # convert 0-based → 1-based
                excel_col = from_col + 1
                label = ''
                for dr in range(-1, 4):
                    for dc in range(-1, 4):
                        text = sheet_cells.get((excel_row + dr, excel_col + dc), '')
                        if len(text.strip()) > 3:
                            label = text.strip()
                            break
                    if label:
                        break

                image_type = classify_image_type(label)

                sheet_images.append({
                    'filename':  img_filename,
                    'path':      img_dest,
                    'label':     label,
                    'type':      image_type,
                    'from_row':  from_row,
                    'from_col':  from_col,
                })

            if sheet_images:
                results.append({
                    'sheet_name': sheet_name,
                    'images': sheet_images,
                })

    return results


# ─────────────────────────────────────────────────────────────────────────────
# Drawing utilities
# ─────────────────────────────────────────────────────────────────────────────

def _draw_rounded_rect_filled(img, x1, y1, x2, y2, color, corner_r=20):
    """Draw a fully-opaque filled rounded rectangle (pill shape) in-place."""
    r = min(corner_r, (x2 - x1) // 2, (y2 - y1) // 2)
    cv2.rectangle(img, (x1 + r, y1), (x2 - r, y2), color, -1)
    cv2.rectangle(img, (x1, y1 + r), (x2, y2 - r), color, -1)
    for cx, cy in [(x1+r, y1+r), (x2-r, y1+r), (x1+r, y2-r), (x2-r, y2-r)]:
        cv2.circle(img, (cx, cy), r, color, -1)


_BADGE_FONT = cv2.FONT_HERSHEY_DUPLEX


def _draw_badge(img, text, bg_color, force_scale=None):
    """Draw a pill-shaped badge, top-right corner of img (in-place).

    If force_scale is given it is used directly; otherwise the scale is derived
    from min(w, h) / 600 so the badge is proportional to the image.
    Pass force_scale from _apply_badge to get identical badges across all images.
    """
    h, w = img.shape[:2]

    # Width-based scale: browser renders every image at the same container width,
    # so badge_display_size = badge_px × (container / img_w) = constant when scale ∝ img_w.
    scale      = force_scale if force_scale is not None else w / 600.0
    font_scale = max(0.75, min(2.2, 1.1 * scale))
    thickness  = max(2, round(font_scale * 2.4))
    pad_x      = max(12, round(20 * scale))   # horizontal inner padding
    pad_y      = max(8,  round(11 * scale))   # vertical inner padding
    corner_r   = max(10, round(18 * scale))   # pill corner radius
    margin     = max(10, round(14 * scale))   # distance from image edges

    (tw, th), baseline = cv2.getTextSize(text, _BADGE_FONT, font_scale, thickness)

    bw = tw + pad_x * 2
    bh = th + pad_y * 2

    x2 = w - margin
    y1 = margin
    x1 = x2 - bw
    y2 = y1 + bh

    # Never overflow left edge
    if x1 < margin:
        x1 = margin
        x2 = x1 + bw

    # Solid pill fill (fully opaque)
    _draw_rounded_rect_filled(img, x1, y1, x2, y2, bg_color, corner_r)

    # Vertically centred text:  ty is the OpenCV baseline coordinate
    # ty = y1 + (bh + th) // 2 - baseline // 2
    tx = x1 + pad_x
    ty = y1 + (bh + th) // 2 - baseline // 2
    cv2.putText(img, text, (tx, ty), _BADGE_FONT, font_scale,
                (255, 255, 255), thickness, cv2.LINE_AA)


def _apply_badge(img_path, text, color, force_scale=None):
    """Read image at img_path, draw badge, overwrite. Used in the second pass."""
    img = cv2.imread(img_path)
    if img is not None:
        _draw_badge(img, text, color, force_scale=force_scale)
        cv2.imwrite(img_path, img)


def _draw_dashed_circle(img, center, radius, color, thickness=2, n_dashes=22):
    """Draw a dashed circle outline."""
    cx, cy = int(center[0]), int(center[1])
    dash_frac = 0.55   # fraction of each segment that is drawn
    for i in range(n_dashes):
        a_start = i       * 360.0 / n_dashes
        a_end   = a_start + 360.0 / n_dashes * dash_frac
        cv2.ellipse(img, (cx, cy), (radius, radius), 0,
                    a_start, a_end, color, thickness, cv2.LINE_AA)


def _get_metric_label(img_type, label_text):
    """Return the short metric name for badge text (e.g. 'SINR', 'RSRP')."""
    t = (label_text or '').lower()
    if img_type == 'quality':
        if 'sinr' in t:   return 'SINR'
        if 'rsrq' in t:   return 'RSRQ'
        if 'rxqual' in t: return 'RxQual'
        return 'Quality'
    if img_type == 'coverage':
        if 'rsrp' in t:  return 'RSRP'
        if 'rscp' in t:  return 'RSCP'
        if 'rxlev' in t: return 'RxLev'
        return 'Coverage'
    if img_type == 'dl_throughput':
        return 'DL Throughput'
    if img_type == 'ul_throughput':
        return 'UL Throughput'
    return 'Signal'


# ─────────────────────────────────────────────────────────────────────────────
# Cross-check math helpers
# ─────────────────────────────────────────────────────────────────────────────

def _hue_dist(h1, h2):
    """Circular hue distance in [0,1] space → result in [0, 0.5]."""
    d = abs(h1 - h2) % 1.0
    return min(d, 1.0 - d)

def _bin_dist(b1, b2, n=36):
    d = abs(b1 - b2) % n
    return min(d, n - d)

def _ang_dist(a, b):
    """Angular distance in degrees → result in [0, 180]."""
    d = abs(a - b) % 360.0
    return min(d, 360.0 - d)

def _angle_from_center(dx, dy):
    """Angle (degrees, 0–360) from site center to pixel.
    Uses atan2(-dy, dx) because image y-axis points downward."""
    return (math.degrees(math.atan2(-dy, dx)) + 360.0) % 360.0

def _circ_mean_deg(angles):
    if not angles:
        return 0.0
    rad = np.deg2rad(angles)
    return float((np.degrees(np.arctan2(np.sin(rad).mean(), np.cos(rad).mean())) + 360.0) % 360.0)

def _density_map(mask, radius):
    """Sum of 1-pixels in a (2r+1)×(2r+1) box centred on each pixel."""
    k = int(2 * radius + 1)
    return cv2.filter2D(mask.astype(np.float32), -1,
                        np.ones((k, k), dtype=np.float32),
                        borderType=cv2.BORDER_CONSTANT)

def _hue_thr_deg(sector_hues, idx):
    dists = [_hue_dist(sector_hues[idx], sector_hues[j]) * 360.0
             for j in range(len(sector_hues)) if j != idx]
    return min(_HUE_CAP_DEG, max(_HUE_FLOOR_DEG, min(dists) * 0.42)) if dists else _HUE_CAP_DEG

def _rgb_thr(prototypes, idx):
    dists = [math.sqrt(sum((float(prototypes[idx][c]) - float(prototypes[j][c]))**2
                           for c in range(3)))
             for j in range(len(prototypes)) if j != idx]
    return min(_RGB_CAP, max(_RGB_FLOOR, min(dists) * 0.42)) if dists else _RGB_CAP


# ─────────────────────────────────────────────────────────────────────────────
# Cross detection — serving cell / PCI / Best Server images
# ─────────────────────────────────────────────────────────────────────────────

def detect_cross(image_path, output_path):
    """
    Proper cross-check algorithm (mirrors analyzer.py):

    1.  Build colorful mask  (S > 0.50, V > 0.35, exclude legend).
    2.  Estimate eNB site center from the densest colored cluster.
    3.  Detect 3 sector hues in a radial fan around the site.
    4.  Sample sector RGB prototypes from the fan.
    5.  Assign every drive-test dot to a sector by color similarity.
    6.  Compute three metrics in angular space:
        - misassigned_ratio  ≥ 0.18 → cross
        - mixed_bin_ratio    ≥ 0.20 → cross
        - max_intrusion_ratio ≥ 0.18 → cross

    Returns (output_path, cross_detected, metrics_dict)
    """
    img = cv2.imread(image_path)
    if img is None:
        return None, False, {'reason': 'image_not_found'}, 'NO CROSS DETECTED', (40, 175, 40)

    h, w = img.shape[:2]

    # ── Resolution scaling ─────────────────────────────────────────────────
    # All distance constants were calibrated at ~1000 px diagonal.
    # Scale proportionally so the algorithm works on any resolution image.
    _diag   = math.sqrt(w * w + h * h)
    _sc     = max(0.5, min(5.0, _diag / 1000.0))

    S_WIN_R  = max(8,   round(_SITE_WIN_R   * _sc))   # site density window
    S_DW     = _SITE_DW / max(1.0, _sc)               # weaker distance penalty on large images
    FAN_IN   = max(4,   round(_FAN_INNER    * _sc))
    FAN_OUT  = max(20,  round(_FAN_OUTER    * _sc))
    FAN_TGT  = max(8,   round(_FAN_TARGET   * _sc))
    FAN_DR   = max(3,   round(_FAN_DEN_R    * _sc))
    FB_WIN   = max(4,   round(_DENSE_WIN_FB * _sc))
    PNT_WR   = max(2,   round(_PNT_WIN_R    * _sc))
    PNT_MR   = max(10,  round(_PNT_MIN_R    * _sc))

    # ── normalized HSV (H: 0→1, S: 0→1, V: 0→1) ─────────────────────────
    hsv_f = cv2.cvtColor(img, cv2.COLOR_BGR2HSV).astype(np.float32)
    hsv_f[:, :, 0] /= 179.0
    hsv_f[:, :, 1] /= 255.0
    hsv_f[:, :, 2] /= 255.0

    rgb_u8 = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)   # (H, W, 3) uint8

    # ── colorful mask ─────────────────────────────────────────────────────
    colorful = ((hsv_f[:, :, 1] >= _CC_SAT_THR) &
                (hsv_f[:, :, 2] >= _CC_VAL_THR)).astype(np.uint8)
    colorful[:int(h * 0.25), :int(w * 0.22)] = 0   # exclude legend

    # ── density maps (pre-compute all needed radii at once) ───────────────
    den_s  = _density_map(colorful, S_WIN_R)   # site centre estimation
    den_f  = _density_map(colorful, FAN_DR)    # fan hue histogram
    den_fb = _density_map(colorful, FB_WIN)    # fallback hue search
    den_p  = _density_map(colorful, PNT_WR)   # point candidates

    # ── 1. Site center ────────────────────────────────────────────────────
    # Expanded search region: nearly the full image (site can be anywhere).
    prior_x = w * _SITE_PRIOR_XR
    prior_y = h * _SITE_PRIOR_YR
    xs0 = int(w * 0.10);  xe0 = int(w * 0.92)
    ys0 = int(h * 0.05);  ye0 = int(h * 0.92)

    region_mask = (colorful[ys0:ye0, xs0:xe0] == 1) & \
                  (den_s[ys0:ye0, xs0:xe0] >= _SITE_MIN_DEN)
    ry, rx = np.where(region_mask)

    if len(rx) == 0:
        # Fallback: image centre-right
        site_x, site_y = prior_x, prior_y
    else:
        abs_x = rx + xs0;  abs_y = ry + ys0
        d2 = (abs_x - prior_x)**2 + (abs_y - prior_y)**2
        dens = den_s[abs_y, abs_x]
        scores = dens - np.sqrt(d2) * S_DW
        order  = np.argsort(-scores)[:_SITE_N_SAMP]
        top_s  = scores[order]
        # Guard: always keep at least the best candidate
        floor  = top_s[0] * _SITE_FLOOR if top_s[0] >= 0 else top_s[0]
        keep   = order[top_s >= floor]
        if len(keep) == 0:
            keep = order[:1]
        w_keep = np.maximum(scores[keep], 1.0)
        site_x = float(np.average(abs_x[keep], weights=w_keep))
        site_y = float(np.average(abs_y[keep], weights=w_keep))

    # ── 2. Sector hues (fan around site) ──────────────────────────────────
    gy, gx = np.mgrid[0:h, 0:w]
    dx_all = gx.astype(np.float32) - site_x
    dy_all = gy.astype(np.float32) - site_y
    r_all  = np.sqrt(dx_all**2 + dy_all**2)

    fan_mask = (colorful == 1) & \
               (r_all >= FAN_IN) & (r_all <= FAN_OUT) & \
               (den_f >= _FAN_MIN_DEN)
    fan_ys, fan_xs = np.where(fan_mask)

    sector_hues = []
    if len(fan_xs) >= 10:
        fan_hues = hsv_f[fan_ys, fan_xs, 0]
        fan_r    = r_all[fan_ys, fan_xs]
        fan_den  = den_f[fan_ys, fan_xs]
        rad_w    = np.maximum(0.2, 1.0 - np.abs(fan_r - FAN_TGT) / FAN_OUT)
        bin_idx  = (fan_hues * 36).astype(int) % 36
        hist     = np.bincount(bin_idx, weights=fan_den * rad_w, minlength=36)

        peaks = []
        for idx in np.argsort(-hist):
            if hist[idx] < _HUE_PEAK_MIN:
                break
            if any(_bin_dist(idx, p) <= 2 for p in peaks):
                continue
            peaks.append(int(idx))
            if len(peaks) == 3:
                break
        sector_hues = [(p + 0.5) / 36.0 for p in peaks]

    # ── fallback: search the entire colorful area ─────────────────────────
    if len(sector_hues) < 3:
        fb_mask = (colorful == 1) & \
                  (gx >= int(w * 0.05)) & (gx < int(w * 0.95)) & \
                  (gy >= int(h * 0.05)) & (gy < int(h * 0.95)) & \
                  (den_fb >= _HUE_PEAK_FB)
        fb_ys, fb_xs = np.where(fb_mask)
        if len(fb_xs) >= 10:
            fb_hues = hsv_f[fb_ys, fb_xs, 0]
            fb_hist = np.bincount((fb_hues * 36).astype(int) % 36, minlength=36)
            peaks   = []
            for idx in np.argsort(-fb_hist):
                if fb_hist[idx] < _HUE_PEAK_FB:
                    break
                if any(_bin_dist(idx, p) <= 2 for p in peaks):
                    continue
                peaks.append(int(idx))
                if len(peaks) == 3:
                    break
            sector_hues = [(p + 0.5) / 36.0 for p in peaks]

    if len(sector_hues) < 3:
        output = img.copy()
        cv2.imwrite(output_path, output)
        return output_path, False, {'reason': 'insufficient_sector_colors'}, \
               'NO CROSS DETECTED', (40, 175, 40)

    # ── 3. Sector RGB prototypes (from the fan) ────────────────────────────
    prototypes = []
    hue_win = _PROTO_HUE_WIN / 360.0
    hf      = hsv_f[:, :, 0]   # (H, W) hue in [0,1]
    proto_base_mask = fan_mask & (r_all >= FAN_IN) & (r_all <= max(FAN_IN + 1, FAN_OUT - 6))
    for sh in sector_hues:
        hd_map  = np.abs(hf - sh)
        hd_map  = np.minimum(hd_map, 1.0 - hd_map)   # circular distance
        proto_mask = proto_base_mask & (hd_map < hue_win)
        py, px = np.where(proto_mask)
        if len(px) == 0:
            py, px = fan_ys, fan_xs       # fallback: all fan pixels
        if len(px) == 0:
            # Last resort: all colorful pixels matching this hue anywhere in image
            wide_mask = (colorful == 1) & (hd_map < hue_win * 2)
            py, px = np.where(wide_mask)
        if len(px) == 0:
            prototypes.append((128, 128, 128))  # neutral fallback
            continue
        samples = rgb_u8[py, px].astype(float)
        prototypes.append(tuple(int(round(v)) for v in samples.mean(axis=0)))

    # ── 4. Assign drive-test dots to sectors ───────────────────────────────
    pt_mask = (colorful == 1) & \
              (den_p >= _PNT_MIN_D) & (den_p <= _PNT_MAX_D) & \
              (r_all >= PNT_MR)
    pt_ys, pt_xs = np.where(pt_mask)

    if len(pt_xs) < _MIN_TOTAL:
        output = img.copy()
        cv2.imwrite(output_path, output)
        return output_path, False, {'reason': 'insufficient_points',
                                    'total': int(len(pt_xs))}, \
               'NO CROSS DETECTED', (40, 175, 40)

    pt_hues = hsv_f[pt_ys, pt_xs, 0]          # (N,)
    pt_rgb  = rgb_u8[pt_ys, pt_xs].astype(float)  # (N, 3)

    sh_arr   = np.array(sector_hues)           # (3,)
    hd       = np.abs(pt_hues[:, None] - sh_arr[None, :])
    hue_dists_deg = np.minimum(hd, 1.0 - hd) * 360.0   # (N, 3)

    proto_arr  = np.array(prototypes, dtype=float)       # (3, 3)
    rgb_dists  = np.sqrt(np.sum(
        (pt_rgb[:, None, :] - proto_arr[None, :, :])**2, axis=2))  # (N, 3)

    # Primary sort: rgb_dist; tiebreak: hue_dist
    combined   = rgb_dists * 1000.0 + hue_dists_deg
    best_sec   = np.argmin(combined, axis=1)  # (N,)

    hue_thrs = [_hue_thr_deg(sector_hues, i) for i in range(3)]
    rgb_thrs = [_rgb_thr(prototypes, i)       for i in range(3)]

    valid = np.zeros(len(pt_ys), dtype=bool)
    for si in range(3):
        m = best_sec == si
        valid |= m & (hue_dists_deg[:, si] <= hue_thrs[si]) & \
                     (rgb_dists[:,    si] <= rgb_thrs[si])

    # ── Connected-component filtering: keep only point-like blobs ─────────────
    # Mirrors is_point_like_component() from analyzer.py:
    #   area 5–90 px, bounding-box span ≤ 14 px on each axis, fill ≥ 0.20
    _COMP_MIN_PX   = 5
    _COMP_MAX_PX   = 90
    _COMP_MAX_SPAN = 14
    _COMP_MIN_FILL = 0.20

    point_angles = []
    for si in range(3):
        m = valid & (best_sec == si)
        # Build a binary mask of this sector's candidate pixels
        sec_mask = np.zeros((h, w), dtype=np.uint8)
        sec_mask[pt_ys[m], pt_xs[m]] = 1

        n_lab, labels, stats, _ = cv2.connectedComponentsWithStats(
            sec_mask, connectivity=8)

        kept_angles = []
        for lab in range(1, n_lab):
            area = int(stats[lab, cv2.CC_STAT_AREA])
            bw   = int(stats[lab, cv2.CC_STAT_WIDTH])
            bh   = int(stats[lab, cv2.CC_STAT_HEIGHT])
            fill = area / (bw * bh) if bw * bh > 0 else 0.0
            if (area < _COMP_MIN_PX or area > _COMP_MAX_PX or
                    bw > _COMP_MAX_SPAN or bh > _COMP_MAX_SPAN or
                    fill < _COMP_MIN_FILL):
                continue
            ys_c, xs_c = np.where(labels == lab)
            angs = np.degrees(np.arctan2(
                -(ys_c.astype(float) - site_y),
                 (xs_c.astype(float) - site_x)
            )) % 360.0
            kept_angles.extend(angs.tolist())

        point_angles.append(kept_angles)

    total_points = sum(len(a) for a in point_angles)
    if total_points < _MIN_TOTAL or any(len(a) < _MIN_PER_CLR for a in point_angles):
        output = img.copy()
        cv2.imwrite(output_path, output)
        return output_path, False, {
            'reason': 'insufficient_points_per_sector',
            'per_sector': [len(a) for a in point_angles],
        }, 'NO CROSS DETECTED', (40, 175, 40)

    # ── 5. Compute evaluation angles & cross metrics ───────────────────────
    eval_angles = [_circ_mean_deg(a) for a in point_angles]
    order = sorted(range(3), key=lambda i: eval_angles[i])
    point_angles = [point_angles[i] for i in order]
    eval_angles  = [eval_angles[i]  for i in order]

    # Sector boundaries = midpoint between adjacent sector centres
    boundaries = []
    n = 3
    for i in range(n):
        left  = eval_angles[i]
        right = eval_angles[(i + 1) % n]
        if i == n - 1 and right < left:
            right += 360.0
        boundaries.append(((left + right) / 2.0) % 360.0)

    # Misassigned ratio
    total_m = mis_m = 0
    for i, angles in enumerate(point_angles):
        for ang in angles:
            total_m += 1
            own_d   = _ang_dist(ang, eval_angles[i])
            other_d = min(_ang_dist(ang, eval_angles[j])
                          for j in range(n) if j != i)
            if other_d + _MISASSIGN_MAR < own_d:
                mis_m += 1
    misassigned = mis_m / total_m if total_m else 1.0

    # Mixed-bin ratio
    bins = [[0] * n for _ in range(36)]
    for i, angles in enumerate(point_angles):
        for ang in angles:
            bins[int(ang // 10) % 36][i] += 1
    mixed_px = tot_px = 0
    for bucket in bins:
        bt = sum(bucket)
        if bt == 0: continue
        tot_px += bt
        if sum(1 for c in bucket if c > 0) >= 2 and max(bucket) / bt < 0.85:
            mixed_px += bt
    mixed_bin = mixed_px / tot_px if tot_px else 1.0

    # Intrusion ratios
    intrusions = []
    for i, angles in enumerate(point_angles):
        start = boundaries[i - 1]
        end   = boundaries[i]
        outside = 0
        for ang in angles:
            ae, as_, aend = ang, start, end
            if aend <= as_: aend += 360.0
            if ae < as_:    ae   += 360.0
            if not (as_ <= ae < aend):
                outside += 1
        intrusions.append(outside / len(angles) if angles else 0.0)
    max_intrusion = max(intrusions)

    cross = (misassigned  >= _MISASSIGN_THR or
             mixed_bin    >= _MIXED_THR     or
             max_intrusion >= _INTRUSION_THR)

    # ── 6. Annotate output image ───────────────────────────────────────────
    output = img.copy()

    # Draw sector direction lines from site center
    sec_colors_bgr = []
    for si in range(n):
        r_, g_, b_ = prototypes[order[si]]
        sec_colors_bgr.append((int(b_), int(g_), int(r_)))  # RGB→BGR

    for i, ang in enumerate(eval_angles):
        rad = math.radians(ang)
        x2  = int(site_x + math.cos(rad)  * 70)
        y2  = int(site_y - math.sin(rad)  * 70)   # note: -sin for image coords
        cv2.line(output, (int(site_x), int(site_y)), (x2, y2),
                 sec_colors_bgr[i], 3, cv2.LINE_AA)

    # Draw sector boundaries
    for bnd in boundaries:
        rad = math.radians(bnd)
        x2  = int(site_x + math.cos(rad) * 62)
        y2  = int(site_y - math.sin(rad) * 62)
        cv2.line(output, (int(site_x), int(site_y)), (x2, y2),
                 (200, 200, 200), 1, cv2.LINE_AA)

    # Draw site center marker
    cv2.circle(output, (int(site_x), int(site_y)), 8,  (255, 255, 255), 2, cv2.LINE_AA)
    cv2.circle(output, (int(site_x), int(site_y)), 3,  (255, 255, 255), -1)

    cv2.imwrite(output_path, output)

    badge_text  = 'CROSS DETECTED'  if cross else 'NO CROSS DETECTED'
    badge_color = (42, 71, 232)      if cross else (40, 175, 40)

    metrics = {
        'cross':              cross,
        'misassigned_ratio':  round(misassigned,    3),
        'mixed_bin_ratio':    round(mixed_bin,       3),
        'max_intrusion_ratio':round(max_intrusion,   3),
        'intrusion_by_sector':intrusions,
        'total_points':       total_points,
        'sector_angles':      [round(a, 1) for a in eval_angles],
        'site_center':        {'x': round(site_x, 1), 'y': round(site_y, 1)},
    }
    return output_path, cross, metrics, badge_text, badge_color


# ─────────────────────────────────────────────────────────────────────────────
# Red zone detection — coverage / quality / throughput images
# ─────────────────────────────────────────────────────────────────────────────

def detect_red_zones(image_path, output_path, min_dot_count=5, padding=25,
                     badge_label='Signal'):
    """
    Detect continuous pure-red drive-test dots and circle degradation zones.

    Strategy:
      1. Strict HSV threshold (H ∈ [0,8] ∪ [172,180], S > 150, V > 150)
         to catch only pure-red dots and exclude orange map markers.
      2. Morphological opening to remove thin text/labels.
      3. Contour filtering by circularity (small) or fill-ratio (large merged blobs).
      4. Graph-based distance clustering (Union-Find) to group nearby dots
         along the drive-test path.
      5. Draw a dashed red circle with semi-transparent fill around each cluster.

    Returns (output_path, zones_list)
    """
    img = cv2.imread(image_path)
    if img is None:
        return None, []

    h_img, w_img = img.shape[:2]
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # Pure red only (excludes orange H≥9)
    lower_r1 = np.array([0,   150, 150])
    upper_r1 = np.array([8,   255, 255])
    lower_r2 = np.array([172, 150, 150])
    upper_r2 = np.array([180, 255, 255])

    red_mask = cv2.bitwise_or(
        cv2.inRange(hsv, lower_r1, upper_r1),
        cv2.inRange(hsv, lower_r2, upper_r2)
    )

    # Exclude legend (top-left corner, ~22% × 30%)
    red_mask[0:int(h_img * 0.30), 0:int(w_img * 0.22)] = 0

    # Exclude footer/logo (bottom ~20% of the image)
    red_mask[int(h_img * 0.80):h_img, 0:w_img] = 0

    # Morphological opening — removes thin text, keeps round blobs
    dot_kernel  = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    dots_only   = cv2.morphologyEx(red_mask, cv2.MORPH_OPEN, dot_kernel)

    contours, _ = cv2.findContours(dots_only, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    dot_centers = []
    dot_mask    = np.zeros_like(red_mask)

    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < 20:
            continue
        perim = cv2.arcLength(cnt, True)
        if perim == 0:
            continue

        circularity     = 4 * np.pi * area / (perim * perim)
        x, y, bw, bh   = cv2.boundingRect(cnt)

        if area <= 2000:
            if circularity < 0.30:
                continue
        else:
            # Large merged trail blob — check fill ratio
            fill = area / (bw * bh) if bw * bh > 0 else 0
            if fill < 0.15:
                continue

        M = cv2.moments(cnt)
        if M['m00'] == 0:
            continue
        cx = int(M['m10'] / M['m00'])
        cy = int(M['m01'] / M['m00'])

        # Large blobs represent several merged dots
        for _ in range(max(1, int(area / 300))):
            dot_centers.append((cx, cy))
        cv2.drawContours(dot_mask, [cnt], -1, 255, -1)

    if not dot_centers:
        shutil.copy(image_path, output_path)
        return output_path, [], f'{badge_label} OK', (40, 175, 40)

    # ── Union-Find distance clustering ────────────────────────────────────
    max_dist = max(w_img, h_img) * 0.08   # 8% of image diagonal
    n = len(dot_centers)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for i in range(n):
        for j in range(i + 1, n):
            dx = dot_centers[i][0] - dot_centers[j][0]
            dy = dot_centers[i][1] - dot_centers[j][1]
            if dx * dx + dy * dy <= max_dist * max_dist:
                union(i, j)

    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(dot_centers[i])

    # ── Draw circles ──────────────────────────────────────────────────────
    output = img.copy()
    zones  = []
    dash_color = (0, 0, 220)   # red outline  (BGR)
    fill_color = (0, 0, 255)   # red fill

    for group_dots in sorted(groups.values(), key=len, reverse=True):
        if len(group_dots) < min_dot_count:
            continue

        xs_g = [d[0] for d in group_dots]
        ys_g = [d[1] for d in group_dots]
        x_min, x_max = min(xs_g), max(xs_g)
        y_min, y_max = min(ys_g), max(ys_g)
        bw = x_max - x_min
        bh = y_max - y_min

        cx     = int((x_min + x_max) / 2)
        cy     = int((y_min + y_max) / 2)
        radius = int(max(bw, bh) / 2) + padding

        zone_id = len(zones) + 1

        # Semi-transparent pink fill
        overlay = output.copy()
        cv2.circle(overlay, (cx, cy), radius, fill_color, -1)
        cv2.addWeighted(overlay, 0.13, output, 0.87, 0, output)

        # Dashed red border (3 px, slightly thicker on small images)
        border_thickness = max(2, int(3 * min(w_img, h_img) / 800))
        _draw_dashed_circle(output, (cx, cy), radius, dash_color,
                            thickness=border_thickness)

        zones.append({
            'id':       zone_id,
            'center_x': cx,
            'center_y': cy,
            'radius':   radius,
            'dots':     len(group_dots),
            'width':    int(bw),
            'height':   int(bh),
        })

    cv2.imwrite(output_path, output)
    badge_text  = f'{badge_label} NOK' if zones else f'{badge_label} OK'
    badge_color = (42, 71, 232)         if zones else (40, 175, 40)
    return output_path, zones, badge_text, badge_color


def _add_clean_badge(output_path, badge_label='Signal'):
    """Add a green OK badge to an already-saved image."""
    img = cv2.imread(output_path)
    if img is not None:
        _draw_badge(img, f'{badge_label} OK', (40, 175, 40))
        cv2.imwrite(output_path, img)


# ─────────────────────────────────────────────────────────────────────────────
# Analysis dispatcher
# ─────────────────────────────────────────────────────────────────────────────

def analyze_image(img_info, session_dir, session_id):
    """Analyze one image based on its classified type."""
    src_path = img_info['path']
    name, ext = os.path.splitext(img_info['filename'])
    out_filename = f'{name}_analyzed{ext}'
    out_path = os.path.join(session_dir, out_filename)

    img_type = img_info['type']
    meta     = IMAGE_TYPE_META.get(img_type, IMAGE_TYPE_META['unknown'])

    entry = {
        'filename':   img_info['filename'],
        'label':      img_info['label'],
        'type':       img_type,
        'type_label': f"{meta['icon']} {meta['label']}",
        'type_color': meta['color'],
        'original':   f'/uploads/{session_id}/{img_info["filename"]}',
        'analyzed':   f'/uploads/{session_id}/{out_filename}',
        'result':     {},
        # Private fields stripped before returning JSON; used for badge second pass
        '_analyzed_path': out_path,
        '_badge_text':    None,
        '_badge_color':   None,
    }

    if img_type == 'geo':
        # No analysis — just copy; no badge needed
        shutil.copy(src_path, out_path)
        entry['result'] = {'kind': 'geo'}

    elif img_type == 'serving_cell':
        _, cross_detected, details, b_text, b_color = detect_cross(src_path, out_path)
        entry['result'] = {
            'kind':           'cross',
            'cross_detected': cross_detected,
            **details,
        }
        entry['_badge_text']  = b_text
        entry['_badge_color'] = b_color

    elif img_type == 'unknown':
        # Unclassified image (logo, legend strip, etc.) — copy only, no analysis
        shutil.copy(src_path, out_path)
        entry['result'] = {'kind': 'unknown'}

    else:
        # coverage / quality / dl_throughput / ul_throughput
        metric = _get_metric_label(img_type, img_info['label'])
        _, zones, b_text, b_color = detect_red_zones(src_path, out_path, badge_label=metric)
        entry['result'] = {
            'kind':  'zones',
            'zones': zones,
        }
        entry['_badge_text']  = b_text
        entry['_badge_color'] = b_color

    return entry


# ─────────────────────────────────────────────────────────────────────────────
# Flask routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route('/favicon.ico')
def favicon():
    return Response(status=204)   # 204 No Content — suppresses the 404 noise


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'Please upload an .xlsx file'}), 400

    session_id  = str(uuid.uuid4())[:8]
    session_dir = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    os.makedirs(session_dir, exist_ok=True)

    filename  = secure_filename(file.filename)
    xlsx_path = os.path.join(session_dir, filename)
    file.save(xlsx_path)

    try:
        sheets = parse_excel_sheets(xlsx_path, session_dir)
    except Exception as e:
        return jsonify({'error': f'Failed to parse Excel file: {e}'}), 400

    if not sheets:
        return jsonify({'error': 'No drive-test images found in this Excel file.'}), 400

    # ── Pass 1: analyze all images (no badges drawn yet) ─────────────────
    response_sheets = []
    all_entries = []   # flat list for badge pass

    for sheet in sheets:
        analyzed = []
        for img_info in sheet['images']:
            try:
                entry = analyze_image(img_info, session_dir, session_id)
            except Exception as e:
                tb = traceback.format_exc()
                print(f'[ERROR] {img_info["filename"]} ({img_info.get("label","")}): {e}\n{tb}')
                entry = {
                    'filename':       img_info['filename'],
                    'label':          img_info.get('label', ''),
                    'type':           'unknown',
                    'type_label':     '⚠ Error',
                    'type_color':     '#ef4444',
                    'original':       f'/uploads/{session_id}/{img_info["filename"]}',
                    'analyzed':       f'/uploads/{session_id}/{img_info["filename"]}',
                    'result':         {'kind': 'error', 'error': str(e)},
                    '_analyzed_path': None,
                    '_badge_text':    None,
                    '_badge_color':   None,
                }
            analyzed.append(entry)
            all_entries.append(entry)
        if analyzed:
            response_sheets.append({
                'sheet_name': sheet['sheet_name'],
                'images':     analyzed,
            })

    # ── Pass 2: draw badges — each image uses its own width-based scale ───
    # Because the browser renders every image at the same container width,
    # scaling badge size ∝ image width guarantees identical badge display size:
    #   badge_display_px = badge_px × (container_w / img_w)
    #                    = (img_w/600 × base) × (container_w / img_w)
    #                    = base × container_w / 600  ← constant for all images
    for e in all_entries:
        p     = e.get('_analyzed_path')
        btext = e.get('_badge_text')
        bcol  = e.get('_badge_color')
        if p and btext and bcol:
            img_tmp = cv2.imread(p)
            if img_tmp is not None:
                _, w_img = img_tmp.shape[:2]
                _draw_badge(img_tmp, btext, bcol, force_scale=w_img / 600.0)
                cv2.imwrite(p, img_tmp)

    # Convert images to base64 (works reliably on Vercel without persistent /tmp)
    for e in all_entries:
        orig_path = e.get('original')
        anal_path = e.get('analyzed')

        # Encode original image
        if orig_path and os.path.exists(orig_path):
            with open(orig_path, 'rb') as f:
                e['original_b64'] = base64.b64encode(f.read()).decode('utf-8')

        # Encode analyzed image
        if anal_path and os.path.exists(anal_path):
            with open(anal_path, 'rb') as f:
                e['analyzed_b64'] = base64.b64encode(f.read()).decode('utf-8')

        # Remove file paths and private keys
        e.pop('original',        None)
        e.pop('analyzed',        None)
        e.pop('_analyzed_path',  None)
        e.pop('_badge_text',     None)
        e.pop('_badge_color',    None)

    return jsonify({
        'filename':   filename,
        'session_id': session_id,
        'sheets':     response_sheets,
    })


@app.route('/uploads/<path:filepath>')
def serve_upload(filepath):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filepath)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, port=port)
