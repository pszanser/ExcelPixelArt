
import io
from dataclasses import dataclass
from typing import List, Optional, Tuple, Dict

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

RGB = Tuple[int, int, int]
Pixel = Optional[RGB]  # None means leave unpainted (transparent)

# ---------------------------- Helpers ----------------------------

EXCEL_GREEN: RGB = (16, 124, 65)      # #107C41
DARK_EXCEL_GREEN: RGB = (8, 94, 47)   # darker accent
WHITE: RGB = (255, 255, 255)
BLACK: RGB = (0, 0, 0)

def to_hex_argb(rgb: RGB) -> str:
    r, g, b = rgb
    return f"FF{r:02X}{g:02X}{b:02X}"

def clamp(x, lo=0, hi=255): 
    return max(lo, min(hi, x))

def lighten(rgb: RGB, amount: float) -> RGB:
    r, g, b = rgb
    return (clamp(int(r + (255 - r) * amount)),
            clamp(int(g + (255 - g) * amount)),
            clamp(int(b + (255 - b) * amount)))

def darken(rgb: RGB, amount: float) -> RGB:
    r, g, b = rgb
    return (clamp(int(r * (1 - amount))),
            clamp(int(g * (1 - amount))),
            clamp(int(b * (1 - amount))))

def color_distance_sq(a: RGB, b: RGB) -> int:
    return (a[0]-b[0])**2 + (a[1]-b[1])**2 + (a[2]-b[2])**2

def load_font(preferred_size: int):
    # Try a few common fonts; fall back to default bitmap
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        try:
            return ImageFont.truetype(p, preferred_size)
        except Exception:
            pass
    return ImageFont.load_default()

# ---------------------------- Core: pixel grids ----------------------------

@dataclass
class MosaicOptions:
    target_width: int = 120
    palette_colors: int = 32
    alpha_threshold: int = 20
    remove_bg_mode: str = "alpha"   # "alpha", "auto-corners", "manual-color"
    manual_bg_color: RGB = (200, 200, 200)
    bg_threshold: int = 22          # color distance threshold for manual/auto background removal

def image_to_pixel_grid(img, opts: MosaicOptions) -> List[List[Pixel]]:
    """Convert an image to a 2D grid of pixels (RGB or None) according to options.
    - If remove_bg_mode != "alpha", attempt to make background transparent by color similarity.
    - The returned grid has width == target_width, height proportional to image.
    """
    im = img.convert("RGBA")
    # Optional: crude background removal by color similarity
    if opts.remove_bg_mode in ("auto-corners", "manual-color"):
        if opts.remove_bg_mode == "auto-corners":
            # average the four corners as background reference
            w, h = im.size
            sample = []
            for (x, y) in [(0,0), (w-1,0), (0,h-1), (w-1,h-1)]:
                sample.append(im.getpixel((x, y))[:3])
            ref = tuple(int(sum(c)/len(c)) for c in zip(*sample))
        else:
            ref = opts.manual_bg_color
        px = im.load()
        w, h = im.size
        thr = opts.bg_threshold ** 2  # squared threshold
        for y in range(h):
            for x in range(w):
                r, g, b, a = px[x, y]
                if a > 0:
                    if color_distance_sq((r, g, b), ref) <= thr:
                        # make it transparent
                        px[x, y] = (r, g, b, 0)

    # Resize while preserving aspect ratio
    target_width = max(10, int(opts.target_width))
    target_height = max(10, int(target_width * im.height / im.width))
    small = im.resize((target_width, target_height), Image.Resampling.LANCZOS)

    # Create a quantized RGB version for palette control (on white background to reduce fringe)
    white_bg = Image.new("RGBA", small.size, (255, 255, 255, 255))
    comp = Image.alpha_composite(white_bg, small)
    quantized = comp.convert("RGB").convert("P", palette=Image.Palette.ADAPTIVE, colors=int(opts.palette_colors)).convert("RGB")

    # Build grid; use alpha to decide transparency
    grid: List[List[Pixel]] = []
    px_small = small.load()
    px_quant = quantized.load()
    for y in range(small.height):
        row: List[Pixel] = []
        for x in range(small.width):
            a = px_small[x, y][3]
            if a <= opts.alpha_threshold:
                row.append(None)
            else:
                row.append(px_quant[x, y])
        grid.append(row)
    return grid

# ---------------------------- Banner generation ----------------------------

@dataclass
class BannerOptions:
    cols: int
    rows: int = 20
    bg_color: RGB = EXCEL_GREEN
    text_color: RGB = WHITE
    accent_color: RGB = WHITE  # for the "40"
    outline_color: RGB = (60, 60, 60)
    show_grid_texture: bool = True
    text_main: str = "40 LAT EXCELA"
    text_sub: str = "Wszystkiego najlepszego! — Piotr"
    style: str = "ExcelBadge"  # ExcelBadge | MinimalXL | Classic

def banner_to_pixel_grid(opts: BannerOptions) -> List[List[RGB]]:
    """Render a banner into a grid of RGB pixels with polished Excel-like styling."""

    def map_to_palette(img: Image.Image, palette: List[RGB]) -> Image.Image:
        # Quantize manually so thin lines snap to the intended palette and avoid washed-out tones.
        px = img.load()
        width, height = img.size
        for y in range(height):
            for x in range(width):
                r, g, b = px[x, y]
                best = palette[0]
                best_dist = float("inf")
                for pr, pg, pb in palette:
                    dist = (r - pr) ** 2 + (g - pg) ** 2 + (b - pb) ** 2
                    if dist < best_dist:
                        best_dist = dist
                        best = (pr, pg, pb)
                px[x, y] = best
        return img

    scale = 12  # draw high-res then downsample for crisp pixel art
    W, H = int(opts.cols * scale), int(opts.rows * scale)
    base = Image.new("RGB", (W, H), opts.bg_color)
    draw = ImageDraw.Draw(base)

    grid_color = lighten(opts.bg_color, 0.18)
    tile_color = darken(opts.bg_color, 0.20)
    shadow_color = darken(opts.bg_color, 0.55)
    palette: List[RGB] = [
        opts.bg_color,
        grid_color,
        tile_color,
        shadow_color,
        opts.text_color,
        opts.accent_color,
        opts.outline_color,
        WHITE,
        BLACK,
    ]

    if opts.show_grid_texture:
        step = max(2, int(6 * scale))
        line_width = max(1, int(0.25 * scale))
        for x in range(0, W, step):
            draw.line([(x, 0), (x, H)], fill=grid_color, width=line_width)
        for y in range(0, H, step):
            draw.line([(0, y), (W, y)], fill=grid_color, width=line_width)

    if opts.style == "ExcelBadge":
        tile_w = max(int(W * 0.18), int(8 * scale))
        margin = max(int(W * 0.02), int(0.6 * scale))
        draw.rectangle([0, 0, tile_w, H], fill=tile_color)

        font_x = load_font(int(H * 0.55))
        bbox_x = draw.textbbox((0, 0), "X", font=font_x)
        x_w = bbox_x[2] - bbox_x[0]
        x_h = bbox_x[3] - bbox_x[1]
        draw.text(
            ((tile_w - x_w) // 2, (H - x_h) // 2),
            "X",
            font=font_x,
            fill=WHITE,
        )

        stroke_big = max(2, int(0.03 * H))
        size_big = int(H * 0.68)
        while True:
            font_big = load_font(size_big)
            bbox_big = draw.textbbox((0, 0), "40", font=font_big, stroke_width=stroke_big)
            width_40 = bbox_big[2] - bbox_big[0]
            if width_40 <= int((W - tile_w - 3 * margin) * 0.5) or size_big <= 18:
                break
            size_big -= 2

        shadow_offset = max(1, int(0.02 * H))
        x_40 = tile_w + margin
        y_40 = int((H * 0.60 - (bbox_big[3] - bbox_big[1])) / 2)
        draw.text(
            (x_40 + shadow_offset, y_40 + shadow_offset),
            "40",
            font=font_big,
            fill=shadow_color,
        )
        draw.text(
            (x_40, y_40),
            "40",
            font=font_big,
            fill=opts.accent_color,
            stroke_width=stroke_big,
            stroke_fill=opts.outline_color,
        )

        stroke_mid = max(1, int(0.02 * H))
        size_mid = int(H * 0.34)
        text_main = opts.text_main if opts.text_main.strip() else "40 LAT EXCELA"
        while True:
            font_mid = load_font(size_mid)
            bbox_mid = draw.textbbox((0, 0), text_main, font=font_mid, stroke_width=stroke_mid)
            if x_40 + width_40 + margin + (bbox_mid[2] - bbox_mid[0]) <= W - margin or size_mid <= 12:
                break
            size_mid -= 1
        draw.text(
            (x_40 + width_40 + margin, int(H * 0.18)),
            text_main,
            font=font_mid,
            fill=opts.text_color,
            stroke_width=stroke_mid,
            stroke_fill=opts.outline_color,
        )

        text_sub = opts.text_sub if opts.text_sub.strip() else "Wszystkiego najlepszego! — Piotr"
        stroke_small = max(1, int(0.018 * H))
        font_small = load_font(int(H * 0.22))
        bbox_small = draw.textbbox((0, 0), text_sub, font=font_small, stroke_width=stroke_small)
        sub_width = bbox_small[2] - bbox_small[0]
        draw.text(
            ((W - sub_width) // 2, int(H * 0.66)),
            text_sub,
            font=font_small,
            fill=opts.text_color,
            stroke_width=stroke_small,
            stroke_fill=opts.outline_color,
        )

    elif opts.style == "MinimalXL":
        stroke_big = max(1, int(0.02 * H))
        text_40 = "40"
        size_big = int(H * 0.68)
        while True:
            font_big = load_font(size_big)
            bbox_big = draw.textbbox((0, 0), text_40, font=font_big, stroke_width=stroke_big)
            width_40 = bbox_big[2] - bbox_big[0]
            if width_40 <= int(W * 0.42) or size_big <= 18:
                break
            size_big -= 2

        x_40 = int(W * 0.05)
        y_40 = int((H - (bbox_big[3] - bbox_big[1])) / 2)
        draw.text(
            (x_40, y_40),
            text_40,
            font=font_big,
            fill=opts.accent_color,
            stroke_width=stroke_big,
            stroke_fill=opts.outline_color,
        )

        text_main = opts.text_main if opts.text_main.strip() else "40 LAT EXCELA"
        stroke_mid = max(1, int(0.018 * H))
        size_mid = int(H * 0.32)
        while True:
            font_mid = load_font(size_mid)
            bbox_mid = draw.textbbox((0, 0), text_main, font=font_mid, stroke_width=stroke_mid)
            if x_40 + width_40 + int(W * 0.03) + (bbox_mid[2] - bbox_mid[0]) <= W - int(W * 0.05) or size_mid <= 14:
                break
            size_mid -= 1
        draw.text(
            (x_40 + width_40 + int(W * 0.03), int(H * 0.2)),
            text_main,
            font=font_mid,
            fill=opts.text_color,
            stroke_width=stroke_mid,
            stroke_fill=opts.outline_color,
        )

        text_sub = opts.text_sub if opts.text_sub.strip() else "Wszystkiego najlepszego! — Piotr"
        font_small = load_font(int(H * 0.20))
        bbox_small = draw.textbbox((0, 0), text_sub, font=font_small)
        sub_width = bbox_small[2] - bbox_small[0]
        draw.text(
            ((W - sub_width) // 2, int(H * 0.70)),
            text_sub,
            font=font_small,
            fill=opts.text_color,
        )

    else:  # Classic
        text_main = opts.text_main if opts.text_main.strip() else "40 LAT EXCELA"
        stroke_mid = max(1, int(0.02 * H))
        font_mid = load_font(int(H * 0.36))
        bbox_mid = draw.textbbox((0, 0), text_main, font=font_mid, stroke_width=stroke_mid)
        mid_width = bbox_mid[2] - bbox_mid[0]
        draw.text(
            ((W - mid_width) // 2, int(H * 0.18)),
            text_main,
            font=font_mid,
            fill=opts.text_color,
            stroke_width=stroke_mid,
            stroke_fill=opts.outline_color,
        )

        text_sub = opts.text_sub if opts.text_sub.strip() else "Wszystkiego najlepszego! — Piotr"
        stroke_small = max(1, int(0.018 * H))
        font_small = load_font(int(H * 0.22))
        bbox_small = draw.textbbox((0, 0), text_sub, font=font_small, stroke_width=stroke_small)
        sub_width = bbox_small[2] - bbox_small[0]
        draw.text(
            ((W - sub_width) // 2, int(H * 0.60)),
            text_sub,
            font=font_small,
            fill=opts.text_color,
            stroke_width=stroke_small,
            stroke_fill=opts.outline_color,
        )

    small = base.resize((opts.cols, opts.rows), Image.Resampling.BOX)
    small = map_to_palette(small, palette)
    return [[small.getpixel((x, y)) for x in range(small.width)] for y in range(small.height)]

# ---------------------------- Excel writer ----------------------------

@dataclass
class CellGeometry:
    col_width: float = 2.15
    row_height: float = 12.15
    margin_top: int = 2       # empty rows before content
    margin_left: int = 2      # empty columns before content
    spacer_rows: int = 2

class FillCache:
    def __init__(self):
        self.cache: Dict[str, PatternFill] = {}

    def get(self, rgb: RGB) -> PatternFill:
        key = to_hex_argb(rgb)
        if key not in self.cache:
            self.cache[key] = PatternFill(start_color=key, end_color=key, fill_type="solid")
        return self.cache[key]

def ensure_square_cells(ws, start_row: int, rows: int, start_col: int, cols: int, geom: CellGeometry):
    for c in range(start_col, start_col + cols):
        ws.column_dimensions[get_column_letter(c)].width = geom.col_width
    for r in range(start_row, start_row + rows):
        ws.row_dimensions[r].height = geom.row_height

def paint_pixels(ws, start_row: int, start_col: int, grid: List[List[Pixel]], fill_cache: FillCache, paint_background: Optional[RGB] = None):
    rows = len(grid)
    cols = len(grid[0]) if rows else 0
    for y in range(rows):
        for x in range(cols):
            rgb = grid[y][x]
            if rgb is None:
                if paint_background is None:
                    continue   # leave unpainted
                rgb = paint_background
            ws.cell(row=start_row + y, column=start_col + x).fill = fill_cache.get(rgb)

def build_workbook(portrait: List[List[Pixel]], banner: List[List[RGB]], geom: CellGeometry, layout: str = "vertical", background: Optional[RGB] = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Excel Pixel Art"
    ws.sheet_view.showGridLines = False

    fill_cache = FillCache()

    start_row = geom.margin_top
    start_col = geom.margin_left

    if layout == "vertical":
        h1 = len(portrait); w1 = len(portrait[0]) if h1 else 0
        h2 = len(banner);   w2 = len(banner[0])   if h2 else 0
        ensure_square_cells(ws, start_row, h1, start_col, w1, geom)
        paint_pixels(ws, start_row, start_col, portrait, fill_cache, paint_background=background)

        for r in range(start_row + h1, start_row + h1 + geom.spacer_rows):
            ws.row_dimensions[r].height = 6

        ensure_square_cells(ws, start_row + h1 + geom.spacer_rows, h2, start_col, w2, geom)
        paint_pixels(ws, start_row + h1 + geom.spacer_rows, start_col, banner, fill_cache, paint_background=background)
    else:
        h1 = len(portrait); w1 = len(portrait[0]) if h1 else 0
        h2 = len(banner);   w2 = len(banner[0])   if h2 else 0

        rows = max(h1, h2)
        ensure_square_cells(ws, start_row, rows, start_col, w1 + geom.spacer_rows + w2, geom)
        paint_pixels(ws, start_row, start_col, portrait, fill_cache, paint_background=background)
        paint_pixels(ws, start_row, start_col + w1 + geom.spacer_rows, banner, fill_cache, paint_background=background)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

# ---------------------------- Streamlit UI ----------------------------

st.set_page_config(page_title="Excel Pixel Art — 40 lat Excela", page_icon=":bar_chart:", layout="wide")
st.title("🎉 Excel Pixel Art — 40 lat Excela")

with st.sidebar:
    st.header("⚙️ Ustawienia")

    target_width = st.slider("Szerokość portretu (kolumny)", min_value=60, max_value=220, value=120, step=2)
    palette_colors = st.slider("Liczba kolorów (paleta)", min_value=8, max_value=64, value=32, step=2)
    alpha_threshold = st.slider("Próg przezroczystości (PNG alpha)", min_value=0, max_value=255, value=20, step=5)

    remove_bg_mode = st.selectbox("Usuwanie tła", ["alpha", "auto-corners", "manual-color"], index=0,
                                  help="alpha — użyj przezroczystości z PNG; auto-corners — próba wycięcia tła po kolorze narożników; manual-color — wskaż kolor tła.")
    manual_bg_hex = st.color_picker("Kolor tła (dla trybu manual)", value="#C8C8C8")
    bg_threshold = st.slider("Tolerancja usuwania tła (większa = agresywniej)", min_value=0, max_value=80, value=22)

    st.markdown("---")
    st.subheader("Baner")
    banner_rows = st.slider("Wysokość banera (wiersze)", min_value=10, max_value=40, value=20, step=1)
    banner_style = st.selectbox("Styl banera", ["ExcelBadge", "MinimalXL", "Classic"], index=0)
    banner_bg_hex = st.color_picker("Kolor tła banera", value="#107C41")
    banner_text_color_hex = st.color_picker("Kolor tekstu", value="#FFFFFF")
    banner_accent_hex = st.color_picker("Kolor akcentu (dla „40”)", value="#FFFFFF")
    show_grid_texture = st.checkbox("Tekstura siatki (subtelna)", value=True)

    text_main_default = "40 LAT EXCELA"
    text_sub_default = "Wszystkiego najlepszego! — Piotr"
    text_main = st.text_input("Nagłówek", value=text_main_default)
    text_sub = st.text_input("Podtytuł", value=text_sub_default)

    st.markdown("---")
    st.subheader("Arkusz")
    layout = st.selectbox("Układ", ["vertical (góra+dół)", "side-by-side (obok siebie)"], index=0)
    col_width = st.slider("Szerokość kolumn (Excel)", min_value=1.2, max_value=4.0, value=2.15, step=0.05)
    row_height = st.slider("Wysokość wierszy (pkt)", min_value=8.0, max_value=22.0, value=12.15, step=0.1)
    paint_background = st.checkbox("Maluj tło (zamiast zostawić białe)", value=False)
    bg_fill_hex = st.color_picker("Kolor tła arkusza (jeśli malujesz tło)", value="#FFFFFF")
    spacer_rows = st.slider("Przerwa między portretem a banerem (wiersze/kolumny)", min_value=0, max_value=6, value=2)

# Utility
def hex_to_rgb(hx: str) -> RGB:
    hx = hx.strip()
    if hx.startswith("#"):
        hx = hx[1:]
    if len(hx) == 3:
        hx = "".join([c*2 for c in hx])
    return (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))

uploaded = st.file_uploader("Wrzuć zdjęcie (PNG z przezroczystością mile widziane, ale JPG też działa)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=False)

col_preview, col_actions = st.columns([2, 1])

if uploaded is not None:
    img = Image.open(uploaded)
    col_preview.image(img, caption="Podgląd wczytanego obrazu", use_container_width=True)

    mo = MosaicOptions(
        target_width=target_width,
        palette_colors=palette_colors,
        alpha_threshold=alpha_threshold,
        remove_bg_mode=remove_bg_mode,
        manual_bg_color=hex_to_rgb(manual_bg_hex),
        bg_threshold=bg_threshold
    )
    portrait_grid = image_to_pixel_grid(img, mo)

    bo = BannerOptions(
        cols=len(portrait_grid[0]),
        rows=banner_rows,
        bg_color=hex_to_rgb(banner_bg_hex),
        text_color=hex_to_rgb(banner_text_color_hex),
        accent_color=hex_to_rgb(banner_accent_hex),
        show_grid_texture=show_grid_texture,
        text_main=text_main if text_main.strip() else "40 LAT EXCELA",
        text_sub=text_sub if text_sub.strip() else "Wszystkiego najlepszego! — Piotr",
        style=banner_style
    )
    banner_grid = banner_to_pixel_grid(bo)

    # Previews
    def grid_to_image(grid: List[List[Pixel]], bg: Optional[RGB]) -> Image.Image:
        h = len(grid)
        w = len(grid[0]) if h else 0
        im = Image.new("RGB", (w, h), bg if bg is not None else WHITE)
        draw = ImageDraw.Draw(im)
        for y in range(h):
            for x in range(w):
                pix = grid[y][x]
                if pix is None:
                    if bg is None:
                        continue
                    pix = bg
                draw.point((x, y), fill=pix)
        return im

    preview_bg = hex_to_rgb(bg_fill_hex) if paint_background else None
    preview_portrait = grid_to_image(portrait_grid, preview_bg).resize((len(portrait_grid[0])*4, len(portrait_grid)*4), Image.NEAREST)
    preview_banner   = grid_to_image(banner_grid, preview_bg).resize((len(banner_grid[0])*4, len(banner_grid)*4), Image.NEAREST)

    with col_preview:
        st.subheader("Podgląd mozaiki (piksele)")
        st.image(preview_portrait, caption="Portret (siatka komórek)", use_container_width=True)
        st.image(preview_banner, caption='Baner „40 lat Excela"', use_container_width=True)

    with col_actions:
        st.subheader("Generowanie pliku")
        geom = CellGeometry(
            col_width=col_width,
            row_height=row_height,
            margin_top=2,
            margin_left=2,
            spacer_rows=spacer_rows
        )
        layout_choice = "vertical" if layout.startswith("vertical") else "horizontal"
        bg_for_sheet = hex_to_rgb(bg_fill_hex) if paint_background else None

        cells_portrait = len(portrait_grid) * len(portrait_grid[0])
        cells_banner = len(banner_grid) * len(banner_grid[0])
        st.caption(f"Szacunkowa liczba komórek: {cells_portrait + cells_banner:,}")

        if st.button("🧩 Generuj plik Excel (.xlsx)"):
            xlsx_bytes = build_workbook(portrait_grid, banner_grid, geom, layout=layout_choice, background=bg_for_sheet)
            st.success("Gotowe! Pobierz plik poniżej:")
            st.download_button("⬇️ Pobierz Excel", data=xlsx_bytes, file_name="Excel_40_lat_pixel_art.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("➡️ Wgraj obraz, aby zacząć. Najlepiej PNG z przezroczystością – wtedy tło w Excelu będzie czyste.")
