
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
    outline_color: RGB = (70, 70, 70)
    show_grid_texture: bool = True
    grid_brightness: float = 0.18  # 0..1 lighten of bg for grid
    text_main: str = "40 LAT EXCELA"
    text_sub: str = "Wszystkiego najlepszego! — Piotr"
    style: str = "XL"  # "XL" for big 40 + small subtitle, "Classic" single line

def banner_to_pixel_grid(opts: BannerOptions) -> List[List[RGB]]:
    """Render a banner into a grid of RGB pixels."""
    scale = 6  # draw high-res then downsample for crisp letters
    W, H = int(opts.cols * scale), int(opts.rows * scale)
    base = Image.new("RGB", (W, H), opts.bg_color)
    drawer = ImageDraw.Draw(base)

    # Optional subtle grid texture to evoke Excel sheets
    if opts.show_grid_texture:
        grid_color = lighten(opts.bg_color, opts.grid_brightness)
        # verticals
        for x in range(0, W, int(6 * scale)):
            drawer.line([(x, 0), (x, H)], fill=grid_color, width=max(1, int(0.3 * scale)))
        # horizontals
        for y in range(0, H, int(6 * scale)):
            drawer.line([(0, y), (W, y)], fill=grid_color, width=max(1, int(0.3 * scale)))

    # Improved typography: large "40" with outline, then "LAT EXCELA" aligned right to it.
    font_big = load_font(int(H * 0.70))
    font_mid = load_font(int(H * 0.35))
    font_small = load_font(int(H * 0.22))

    if opts.style == "XL":
        text_40 = "40"
        bbox_40 = drawer.textbbox((0, 0), text_40, font=font_big, stroke_width=max(1, int(0.04*H)))
        w40, h40 = bbox_40[2] - bbox_40[0], bbox_40[3] - bbox_40[1]

        x40 = int(W * 0.05)
        y40 = int((H * 0.60 - h40) / 2)

        drawer.text((x40, y40), text_40, font=font_big,
                    fill=opts.accent_color,
                    stroke_width=max(2, int(0.03 * H)),
                    stroke_fill=opts.outline_color)

        text_main = "LAT EXCELA"
        bbox_main = drawer.textbbox((0,0), text_main, font=font_mid, stroke_width=max(1, int(0.025*H)))
        wmain, hmain = bbox_main[2] - bbox_main[0], bbox_main[3] - bbox_main[1]
        xmain = x40 + w40 + int(W * 0.03)
        ymain = int(H * 0.18)
        drawer.text((xmain, ymain), text_main, font=font_mid,
                    fill=opts.text_color,
                    stroke_width=max(1, int(0.025*H)),
                    stroke_fill=opts.outline_color)

        text_sub = opts.text_sub
        bbox_sub = drawer.textbbox((0,0), text_sub, font=font_small, stroke_width=max(1, int(0.02*H)))
        wsub = bbox_sub[2] - bbox_sub[0]
        xsub = int((W - wsub) / 2)
        ysub = int(H * 0.64)
        drawer.text((xsub, ysub), text_sub, font=font_small,
                    fill=opts.text_color,
                    stroke_width=max(1, int(0.02*H)),
                    stroke_fill=opts.outline_color)
    else:
        # Classic single line headline + subline
        text_main = opts.text_main
        bbox_main = drawer.textbbox((0,0), text_main, font=font_mid, stroke_width=max(1, int(0.02*H)))
        wmain, hmain = bbox_main[2] - bbox_main[0], bbox_main[3] - bbox_main[1]
        xmain = int((W - wmain) / 2); ymain = int(H * 0.18)
        drawer.text((xmain, ymain), text_main, font=font_mid,
                    fill=opts.text_color,
                    stroke_width=max(1, int(0.02*H)),
                    stroke_fill=opts.outline_color)

        text_sub = opts.text_sub
        bbox_sub = drawer.textbbox((0,0), text_sub, font=font_small, stroke_width=max(1, int(0.02*H)))
        wsub = bbox_sub[2] - bbox_sub[0]
        xsub = int((W - wsub) / 2); ysub = int(H * 0.60)
        drawer.text((xsub, ysub), text_sub, font=font_small,
                    fill=opts.text_color,
                    stroke_width=max(1, int(0.02*H)),
                    stroke_fill=opts.outline_color)

    small = base.resize((opts.cols, opts.rows), Image.Resampling.LANCZOS)
    pixels = [[small.getpixel((x, y)) for x in range(small.width)] for y in range(small.height)]
    return pixels

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
    banner_style = st.selectbox("Styl banera", ["XL", "Classic"], index=0)
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
