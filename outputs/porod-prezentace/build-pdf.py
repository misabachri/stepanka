from pathlib import Path
from textwrap import wrap

from PIL import Image
from reportlab.lib.colors import HexColor, Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent / "porod-neni-jen-medicina.pdf"
PHOTO = ROOT / "foto prezentace"
W, H = 1280, 720

C = {
    "bg": "#ffffff",
    "bg_soft": "#fffbfb",
    "bg_warm": "#fff5f0",
    "rose": "#dec0b2",
    "taupe": "#b79c8b",
    "accent": "#654334",
    "cta": "#911c0b",
    "text": "#361b0f",
    "white": "#ffffff",
}

RADIUS = 7


def register_fonts():
    fonts = {
        "Head": "/System/Library/Fonts/Supplemental/Georgia.ttf",
        "Head-Bold": "/System/Library/Fonts/Supplemental/Georgia Bold.ttf",
        "Body": "/System/Library/Fonts/Supplemental/Arial.ttf",
        "Body-Bold": "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    }
    for name, file in fonts.items():
        pdfmetrics.registerFont(TTFont(name, file))


def cover_box(img_path, x, y, w, h):
    with Image.open(img_path) as im:
        iw, ih = im.size
    ir, fr = iw / ih, w / h
    if ir > fr:
        visible_w = ih * fr
        sx = (iw - visible_w) / 2
        crop = (sx, 0, sx + visible_w, ih)
    else:
        visible_h = iw / fr
        sy = (ih - visible_h) / 2
        crop = (0, sy, iw, sy + visible_h)

    tmp_dir = Path(__file__).resolve().parents[2] / "tmp/slides/porod-prezentace/pdf-images"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    out = tmp_dir / f"{img_path.stem}-{int(w)}x{int(h)}.jpg"
    with Image.open(img_path) as im:
        im = im.convert("RGB").crop(tuple(map(int, crop))).resize((int(w), int(h)), Image.LANCZOS)
        im.save(out, quality=92)
    return out


def contain_box(img_path, w, h):
    tmp_dir = Path(__file__).resolve().parents[2] / "tmp/slides/porod-prezentace/pdf-images"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    out = tmp_dir / f"{img_path.stem}-contain-{int(w)}x{int(h)}.jpg"
    with Image.open(img_path) as im:
        im = im.convert("RGB")
        iw, ih = im.size
        scale = min(w / iw, h / ih)
        nw, nh = int(iw * scale), int(ih * scale)
        resized = im.resize((nw, nh), Image.LANCZOS)
        base = Image.new("RGB", (int(w), int(h)), C["bg_warm"])
        base.paste(resized, ((int(w) - nw) // 2, (int(h) - nh) // 2))
        base.save(out, quality=92)
    return out


def fill(c, color, alpha=1):
    c.setFillColor(HexColor(color) if alpha >= 1 else Color(*HexColor(color).rgb(), alpha=alpha))


def text(c, value, x, y, size=24, font="Body", color="text", leading=None, max_chars=None):
    fill(c, C[color])
    c.setFont(font, size)
    leading = leading or size * 1.18
    lines = []
    for part in str(value).split("\n"):
        lines.extend(wrap(part, width=max_chars or 42) or [""])
    for i, line in enumerate(lines):
        c.drawString(x, y - i * leading, line)
    return y - len(lines) * leading


def title(c, value, x, y, size=52, max_chars=18):
    return text(c, value, x, y, size=size, font="Head-Bold", color="text", leading=size * 1.18, max_chars=max_chars)


def bullet_list(c, items, x, y, size=23, max_chars=42):
    fill(c, C["text"])
    c.setFont("Body", size)
    leading = size * 1.45
    cur = y
    for item in items:
        lines = wrap(item, width=max_chars) or [item]
        fill(c, C["cta"])
        c.circle(x + 4, cur - 7, 4, fill=1, stroke=0)
        fill(c, C["text"])
        c.drawString(x + 22, cur, lines[0])
        for line in lines[1:]:
            cur -= leading * .86
            c.drawString(x + 22, cur, line)
        cur -= leading
    return cur


def photo(c, image, x, y, w, h):
    img = cover_box(PHOTO / image, x, y, w, h)
    c.drawImage(str(img), x, y, w, h, mask="auto")


def photo_contain(c, image, x, y, w, h):
    img = contain_box(PHOTO / image, w, h)
    c.drawImage(str(img), x, y, w, h, mask="auto")


def card(c, x, y, w, h, color="white", alpha=0.82):
    fill(c, C[color], alpha)
    c.roundRect(x, y, w, h, RADIUS, fill=1, stroke=0)


def separator(c, x, y, w=168, h=2):
    fill(c, C["cta"])
    c.roundRect(x, y, w, h, RADIUS, fill=1, stroke=0)


def quote_note(c, value, x, y, w=452, h=78, size=22, max_chars=38):
    card(c, x, y, w, h, "bg_warm", .92)
    text(c, "“", x + 16, y + h - 8, 54, "Head-Bold", "rose", max_chars=2)
    text(c, value, x + 58, y + h - 29, size, "Head-Bold", "cta", max_chars=max_chars)


slides = [
    ("title", "1.jpg", "Porod není jen medicína", [], "aneb proč na porodu záleží", "Bc. Štěpánka Trappová", "komunitní porodní asistentka"),
    ("content", "2.jpg", "Kdo jsem a co dělám", ["Doprovázím ženy těhotenstvím, porodem i po porodu.", "Nabízím předporodní přípravu, doprovod k porodu a poporodní péči."], "Jsem porodní asistentka.", None, "Moje práce není jen o zdravotní péči. Je hlavně o podpoře, důvěře a bezpečí."),
    ("statement", "3.jpg", "Proč právě porod?", ["Je to silný životní moment.", "Zážitek, který si ženy pamatují celý život.", "Může být posilující, nebo naopak zraňující."], "Protože porod není jen „lékařská událost“.", None, None),
    ("content", "4.jpg", "Jak si ženy porod představují", ["žena leží na zádech", "hodně zásahů a léků", "rychlý, dramatický průběh", "často pasivní role ženy"], "Podle filmů a médií:", None, None),
    ("content", "5.jpg", "Realita v systému", ["méně prostoru pro individuální přístup", "časový tlak", "zavedené postupy", "omezený prostor pro vytvoření důvěry"], None, None, "Ne proto, že by někdo chtěl ublížit. Ale protože systém má svoje limity."),
    ("content", "6.jpg", "Porod může vypadat i jinak", ["gauč", "voda", "porodní stolička", "postel"], "Různé polohy, více svobody pohybu a větší zapojení ženy.", None, "Každé ženě vyhovuje něco jiného."),
    ("statement", "7.jpg", "Neexistuje jeden „správný“ porod", ["každá žena je jiná", "každé tělo je jiné", "každá zkušenost je jiná"], None, None, "Důležité je, aby se žena cítila bezpečně, respektovaně a vyslyšeně."),
    ("statement", "8.jpg", "Proč na porodu záleží", ["porod si žena pamatuje celý život", "většinou ho zažije jen 1–2×"], None, None, "Není to „jen jeden den“."),
    ("content", "jakporod.jpg", "Jak porod ovlivňuje ženu", ["sebevědomí", "vztah k sobě", "vztah k dítěti", "začátek kojení"], None, None, "Může ji posílit... nebo naopak oslabit."),
    ("content", "10a.jpg", "Proč dělám to, co dělám", ["podporu", "respekt", "individuální přístup", "pochopení", "nesouzení"], "Protože věřím, že porod může být posilující zážitek a ženy mají v sobě obrovskou sílu.", None, None),
    ("content", "11 (1).jpg", "Co dělá rozdíl", ["když žena není „jen další pacientka“", "když má kolem sebe známé lidi", "když má kontinuální péči"], None, None, "Důvěra = klíč k dobrému zážitku."),
    ("content", "come.jpg", "Co mě ovlivnilo", ["zkušenost z Německa", "vlastní porod", "ukázalo mi to, že věci mohou fungovat i jinak"], None, None, None),
    ("statement", "13.jpg", "Na porod nemusíte být sama", ["můžete mít podporu", "můžete mít informace", "můžete mít někoho na své straně"], None, None, "A přesně to nabízím."),
    ("contact", "14.jpg", "Pokud vás to oslovilo...", ["Bc. Štěpánka Trappová", "+420 605 074 332", "stepanka@trappea.cz", "www.trappea.cz"], "Ráda vás podpořím. Můžete se na mě obrátit.", None, None),
]


def draw_content(c, idx, image, heading, items, lead, note):
    fill(c, C["bg_soft"])
    c.rect(0, 0, W, H, fill=1, stroke=0)
    reverse = idx % 2 == 0
    tx, ix = (650, 70) if reverse else (92, 735)
    card(c, tx - 30, 86, 520, 548)
    photo(c, image, ix, 70, 480, 585)
    y = 548
    y = title(c, heading, tx, y, 42, 19)
    separator(c, tx, y - 2)
    y -= 34
    if lead:
        y = text(c, lead, tx, y, 23, "Body-Bold", "accent", max_chars=42) - 18
    if items:
        y = bullet_list(c, items, tx, y, 22, 42) - 8
    if note:
        quote_note(c, note, tx, max(86, y - 80))
    text(c, str(idx).zfill(2), 1130, 38, 14, "Head-Bold", "taupe")


def draw():
    register_fonts()
    c = canvas.Canvas(str(OUT), pagesize=(W, H), pageCompression=1)

    for idx, (kind, image, heading, items, lead, kicker, note) in enumerate(slides, start=1):
        if kind == "title":
            fill(c, C["bg_warm"])
            c.rect(0, 0, W, H, fill=1, stroke=0)
            fill(c, C["bg_soft"])
            c.rect(0, 0, 642, H, fill=1, stroke=0)
            photo(c, image, 705, 44, 496, 632)
            text(c, kicker, 76, 592, 21, "Body-Bold", "accent")
            text(c, note, 76, 552, 18, "Body", "accent")
            title(c, heading, 76, 462, 70, 15)
            separator(c, 76, 286)
            text(c, lead, 76, 248, 38, "Head-Bold", "accent", max_chars=28)
        elif kind == "statement":
            photo(c, image, 0, 0, W, H)
            fill(c, C["bg_warm"], .78)
            c.rect(0, 0, W, H, fill=1, stroke=0)
            card(c, 92, 92, 700, 500, "white", .92)
            title(c, heading, 136, 500, 54, 20)
            separator(c, 146, 386)
            y = 350
            if lead:
                y = text(c, lead, 146, y, 24, "Body-Bold", "accent", max_chars=43) - 22
            y = bullet_list(c, items, 146, y, 25, 44)
            if note:
                quote_note(c, note, 146, 150, 500, 84, 31, 30)
            text(c, str(idx).zfill(2), 1130, 38, 14, "Head-Bold", "white")
        elif kind == "image":
            fill(c, C["bg_warm"])
            c.rect(0, 0, W, H, fill=1, stroke=0)
            photo_contain(c, image, 410, 42, 460, 636)
        elif kind == "contact":
            fill(c, C["bg_warm"])
            c.rect(0, 0, W, H, fill=1, stroke=0)
            card(c, 92, 90, 590, 540)
            photo(c, image, 746, 168, 384, 384)
            title(c, heading, 136, 538, 48, 22)
            separator(c, 140, 438)
            text(c, lead, 140, 406, 24, "Body-Bold", "accent", max_chars=40)
            fill(c, C["cta"])
            c.roundRect(140, 346, 80, 4, RADIUS, fill=1, stroke=0)
            y = 294
            for i, item in enumerate(items):
                y = text(c, item, 140, y, 24 if i == 0 else 22, "Head-Bold" if i == 0 else "Body", "text" if i == 0 else "accent") - 8
            text(c, str(idx).zfill(2), 1130, 38, 14, "Head-Bold", "taupe")
        else:
            draw_content(c, idx, image, heading, items, lead, note)
        c.showPage()

    c.save()
    print(OUT)


if __name__ == "__main__":
    draw()
