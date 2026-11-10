import asyncio
import io
import logging
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml.ns import qn
from dotenv import load_dotenv
from pdf2docx import Converter
from PIL import Image, ImageOps
from pypdf import PdfWriter
from telegram import (
    BotCommand,
    InlineKeyboardButton as Btn,
    InlineKeyboardMarkup as Markup,
    InputMediaDocument,
    Update,
)
from telegram.constants import ChatAction, ParseMode
from telegram.error import TelegramError
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

try:
    import pytesseract
except ImportError:  # OCR is optional
    pytesseract = None

load_dotenv()
logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO
)
# បិទ Log របស់ httpx ដើម្បីកុំឱ្យបង្ហាញ Bot Token ក្នុង Terminal
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("converter")

HTML = ParseMode.HTML
MAX_SIZE = 20 * 1024 * 1024  # Telegram bot download limit
MAX_PAGES = 30  # max pages sent back for PDF -> JPG / split-all
MAX_FILES = 50  # files collected per job
AUTO_DELAY = 2.5  # seconds of silence after an album before auto-creating the PDF
PAGE_DPI = 200  # quality of generated PDF pages (A4 at 200 dpi)
DEFAULT_MARGIN_MM = 5  # white border around images in the PDF (change in ⚙️ settings)
KHMER_FONT = os.environ.get("KHMER_FONT", "Khmer OS Siemreap")
OCR_LANG = os.environ.get("OCR_LANG", "khm+eng")
OCR_MAX_PAGES = 30  # OCR is slow, limit pages
TASKS: set = set()  # keep references to background tasks

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}
WORD_EXT = {".doc", ".docx", ".odt", ".rtf"}
WORD_MIMES = {
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.oasis.opendocument.text",
    "application/rtf",
}

# mode -> button label
TOOLS = {
    "img2pdf": "🖼 JPG → PDF",
    "pdf2img": "📄 PDF → JPG",
    "word2pdf": "📘 Word → PDF",
    "pdf2word": "📝 PDF → Word",
    "merge": "📎 បញ្ចូល PDF",
    "split": "✂️ បំបែក PDF",
    "compress": "🗜 បង្រួម PDF",
    "rotate": "🔄 បង្វិល PDF",
    "pagenum": "🔢 ដាក់លេខទំព័រ",
    "protect": "🔒 ដាក់លេខសម្ងាត់",
    "unlock": "🔓 ដោះលេខសម្ងាត់",
}
MULTI = ("img2pdf", "merge")  # modes that collect several files
INTRO = {
    "pdf2img": f"ផ្ញើឯកសារ PDF មក 📤\nខ្ញុំនឹងបំប្លែងគ្រប់ទំព័រជារូប JPG (អតិបរមា {MAX_PAGES} ទំព័រ)",
    "word2pdf": "ផ្ញើឯកសារ Word (.docx / .doc) មក 📤\nខ្ញុំនឹងបំប្លែងជា PDF ដោយរក្សាទម្រង់អក្សរខ្មែរនៅដដែល",
    "pdf2word": "ផ្ញើឯកសារ PDF មក 📤\nខ្ញុំនឹងបំប្លែងជាឯកសារ Word (.docx) គាំទ្រអក្សរខ្មែរ",
    "split": "ផ្ញើឯកសារ PDF ដែលចង់បំបែក 📤",
    "compress": "ផ្ញើឯកសារ PDF ដែលចង់បង្រួមទំហំ 📤",
    "rotate": "ផ្ញើឯកសារ PDF ដែលចង់បង្វិល 📤",
    "pagenum": "ផ្ញើឯកសារ PDF ដែលចង់ដាក់លេខទំព័រ 📤",
    "protect": "ផ្ញើឯកសារ PDF ដែលចង់ដាក់លេខសម្ងាត់ 📤",
    "unlock": "ផ្ញើឯកសារ PDF ដែលមានលេខសម្ងាត់ 📤",
}
LEVELS = {"high": (150, 80), "mid": (110, 65), "low": (72, 50)}  # (dpi, jpeg quality)

WELCOME = (
    "✨ <b>Fast Convert</b> ✨\n"
    "ឧបករណ៍បំប្លែង និងកែសម្រួលឯកសារ PDF / Word ឆាប់រហ័ស (គាំទ្រអក្សរខ្មែរ 🇰🇭)\n\n"
    "💡 ផ្ញើរូបភាព, Word ឬ PDF មកផ្ទាល់ក៏បាន ឬជ្រើសរើសឧបករណ៍ខាងក្រោម 👇"
)
HELP = (
    "ℹ️ <b>របៀបប្រើ</b>\n\n"
    "🖼 <b>JPG → PDF</b>: ផ្ញើរូបច្រើនក្នុងពេលតែមួយ ខ្ញុំបង្កើត PDF ឱ្យដោយស្វ័យប្រវត្តិ\n"
    "📄 <b>PDF → JPG</b>: បំប្លែងទំព័រនីមួយៗជារូបភាព\n"
    "📘 <b>Word → PDF</b>: បំប្លែងឯកសារ Word (.doc/.docx) ទៅជា PDF (គាំទ្រអក្សរខ្មែរ)\n"
    "📝 <b>PDF → Word</b>: បំប្លែងឯកសារ PDF ទៅជា Word (.docx) អាចកែអក្សរខ្មែរបាន\n"
    "📎 <b>បញ្ចូល PDF</b>: ផ្ញើ PDF ច្រើនតាមលំដាប់\n"
    "✂️ <b>បំបែក PDF</b>: ទាញយកទំព័រដែលចង់បាន (ឧ. 1-3,5)\n"
    "🗜 <b>បង្រួម PDF</b>: បន្ថយទំហំឯកសារ\n"
    "🔄 <b>បង្វិល PDF</b>: បង្វិលទំព័រ 90° / 180° / 270°\n"
    "🔢 <b>ដាក់លេខទំព័រ</b>: ដាក់លេខនៅខាងក្រោមទំព័រ\n"
    "🔒/🔓 <b>លេខសម្ងាត់</b>: ដាក់ ឬដោះលេខសម្ងាត់ PDF\n"
    "⚙️ <b>ការកំណត់</b>: ទំហំទំព័រ និងគែមនៃ JPG → PDF\n\n"
    f"📏 ទំហំអតិបរមា {MAX_SIZE // (1024 * 1024)}MB ក្នុងមួយឯកសារ"
)


class UserError(Exception):
    """Problem the user can fix (bad page range, wrong password...)."""


# ---------- keyboards & texts ----------

def main_menu() -> Markup:
    def b(m):
        return Btn(TOOLS[m], callback_data=f"mode:{m}")

    return Markup(
        [
            [b("img2pdf"), b("pdf2img")],
            [b("word2pdf"), b("pdf2word")],
            [b("merge"), b("split")],
            [b("compress"), b("rotate")],
            [b("pagenum"), b("protect")],
            [b("unlock"), Btn("⚙️ ការកំណត់", callback_data="settings")],
            [Btn("ℹ️ ជំនួយ", callback_data="help")],
        ]
    )


def back_kb() -> Markup:
    return Markup([[Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")]])


def collect_kb(n: int) -> Markup:
    if n == 0:
        return back_kb()
    return Markup(
        [
            [Btn("✅ បង្កើត PDF", callback_data="done")],
            [
                Btn("🗑 សម្អាត", callback_data="clear"),
                Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu"),
            ],
        ]
    )


def collect_text(mode: str, n: int) -> str:
    title = f"<b>{TOOLS[mode]}</b>"
    noun = "រូបភាព" if mode == "img2pdf" else "PDF"
    if n == 0:
        return (
            f"{title}\n\nផ្ញើ{noun}មកខ្ញុំ 📤\n"
            f"💡 ជ្រើសរើស{noun}ច្រើនក្នុងពេលតែមួយ ខ្ញុំនឹងបង្កើត PDF ដោយស្វ័យប្រវត្តិ"
        )
    return (
        f"{title}\n\n📥 បានទទួល <b>{n}</b> {noun}\n"
        "ផ្ញើបន្ថែម ឬចុច <b>បង្កើត PDF</b> ដើម្បីបញ្ចប់"
    )


def intro_text(mode: str) -> str:
    if mode in MULTI:
        return collect_text(mode, 0)
    return f"<b>{TOOLS[mode]}</b>\n\n{INTRO[mode]}"


def pdf_choice_kb(encrypted: bool) -> Markup:
    if encrypted:
        return Markup(
            [
                [Btn(TOOLS["unlock"], callback_data="use:unlock")],
                [Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")],
            ]
        )

    def b(m):
        return Btn(TOOLS[m], callback_data=f"use:{m}")

    return Markup(
        [
            [b("pdf2word"), b("pdf2img")],
            [b("merge"), b("split")],
            [b("compress"), b("rotate")],
            [b("pagenum"), b("protect")],
            [Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")],
        ]
    )


def settings_view(cfg: dict) -> tuple[str, Markup]:
    page_name = "A4" if cfg["page"] == "a4" else "តាមទំហំរូប"
    text = (
        "⚙️ <b>ការកំណត់ «JPG → PDF»</b>\n\n"
        f"📐 ទំហំទំព័រ: <b>{page_name}</b>\n"
        f"🔲 គែមសជុំវិញរូប: <b>{cfg['margin']} មម</b>"
    )

    def t(ok: bool) -> str:
        return "✅ " if ok else ""

    kb = Markup(
        [
            [
                Btn(f"{t(cfg['page'] == 'a4')}📐 A4", callback_data="set:page:a4"),
                Btn(f"{t(cfg['page'] == 'fit')}📐 តាមទំហំរូប", callback_data="set:page:fit"),
            ],
            [
                Btn(f"{t(cfg['margin'] == 0)}គ្មានគែម", callback_data="set:margin:0"),
                Btn(f"{t(cfg['margin'] == 5)}តូច 5មម", callback_data="set:margin:5"),
                Btn(f"{t(cfg['margin'] == 10)}មធ្យម 10មម", callback_data="set:margin:10"),
            ],
            [Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")],
        ]
    )
    return text, kb


def fmt_size(n: int) -> str:
    return f"{n / 1024:.0f} KB" if n < 1024 * 1024 else f"{n / 1024 / 1024:.1f} MB"


# ---------- per-user state ----------

def state(ctx: ContextTypes.DEFAULT_TYPE) -> dict:
    return ctx.user_data.setdefault(
        "s",
        {
            "mode": None,
            "files": [],
            "status_id": None,
            "pdf": None,
            "name": None,
            "info": None,
            "await": None,
            "seq": 0,
            "lock": asyncio.Lock(),
        },
    )


def cfg_of(ctx: ContextTypes.DEFAULT_TYPE) -> dict:
    return ctx.user_data.setdefault("cfg", {"margin": DEFAULT_MARGIN_MM, "page": "a4"})


def reset(ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    s["seq"] += 1
    s.update(mode=None, files=[], status_id=None, pdf=None, name=None, info=None)
    s["await"] = None


async def safe_edit(q, text: str, kb: Markup | None = None):
    try:
        await q.edit_message_text(text, reply_markup=kb, parse_mode=HTML)
    except TelegramError:
        pass


# ---------- Khmer & Word/PDF helpers ----------

_KHMER_RE = re.compile(r"[\u1780-\u17FF\u19E0-\u19FF]")


def _apply_khmer_font_to_runs(paragraphs, font_name: str):
    for p in paragraphs:
        for run in p.runs:
            if run.text and _KHMER_RE.search(run.text):
                run.font.name = font_name
                rPr = run._element.get_or_add_rPr()
                rFonts = rPr.get_or_add_rFonts()
                rFonts.set(qn("w:ascii"), font_name)
                rFonts.set(qn("w:hAnsi"), font_name)
                rFonts.set(qn("w:cs"), font_name)
                rFonts.set(qn("w:eastAsia"), font_name)
                # complex-script font size must match the normal size
                sz = rPr.find(qn("w:sz"))
                if sz is not None and rPr.find(qn("w:szCs")) is None:
                    szCs = rPr.makeelement(qn("w:szCs"), {qn("w:val"): sz.get(qn("w:val"))})
                    rPr.append(szCs)


def enforce_khmer_docx_fonts(docx_path: Path, font_name: str = KHMER_FONT):
    doc = Document(str(docx_path))
    _apply_khmer_font_to_runs(doc.paragraphs, font_name)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _apply_khmer_font_to_runs(cell.paragraphs, font_name)
    for section in doc.sections:
        _apply_khmer_font_to_runs(section.header.paragraphs, font_name)
        _apply_khmer_font_to_runs(section.footer.paragraphs, font_name)
    doc.save(str(docx_path))


def khmer_text_broken(text: str) -> bool:
    """ទាយថាអក្សរខ្មែរដែលទាញបានខូច ឬអត់ (ឬ PDF ជារូបស្កេន)"""
    if not text.strip() or "\ufffd" in text:
        return True
    bad = (
        len(re.findall(r"\u17D2(?![\u1780-\u17A2])", text))  # ជើង ្ មិនមានព្យញ្ជនៈតាម
        + len(re.findall(r"(?:^|\s)[\u17B6-\u17C5]", text))  # ស្រៈនៅដើមពាក្យ
    )
    return bad >= 3


def ocr_pdf_to_docx(data: bytes, out_path: Path):
    if pytesseract is None:
        raise UserError(
            "PDF នេះមានអក្សរខ្មែរដែលទាញមិនបាន ហើយម៉ាស៊ីនមេមិនទាន់ដំឡើង OCR (pytesseract) ទេ"
        )
    doc = pymupdf.open(stream=data, filetype="pdf")
    word = Document()
    try:
        for i, page in enumerate(doc):
            if i >= OCR_MAX_PAGES:
                break
            if i:
                word.add_page_break()
            img = Image.open(io.BytesIO(page.get_pixmap(dpi=300).tobytes("png")))
            try:
                text = pytesseract.image_to_string(img, lang=OCR_LANG)
            except pytesseract.TesseractError as e:
                log.error("tesseract failed: %s", e)
                raise UserError(
                    "OCR មិនដំណើរការ៖ សូមដំឡើងទិន្នន័យភាសាខ្មែរ (tesseract-ocr-khm)"
                )
            for line in text.splitlines():
                word.add_paragraph(line)
    finally:
        doc.close()
    word.save(str(out_path))


def find_soffice() -> str:
    """ស្វែងរក LibreOffice ទាំងលើ macOS, Linux និង Windows"""
    for cmd in ("soffice", "libreoffice"):
        path = shutil.which(cmd)
        if path:
            return path
    extra_paths = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
        "/opt/local/bin/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in extra_paths:
        if os.path.isfile(p):
            return p
    raise UserError(
        "ម៉ាស៊ីនមេមិនទាន់បានដំឡើង LibreOffice សម្រាប់បំប្លែង Word → PDF ទេ។"
    )


# ---------- PDF / image / Word processing (CPU work, run in threads) ----------

def mm(v: float) -> int:
    return round(v / 25.4 * PAGE_DPI)


def to_rgb(data: bytes) -> Image.Image:
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data)))
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.split()[-1])
        return bg
    return img.convert("RGB")


def to_page(img: Image.Image, margin_mm: float, page: str) -> Image.Image:
    m = mm(margin_mm)
    if page == "fit":
        scale = min(1.0, mm(297) / max(img.size))
        if scale < 1:
            img = img.resize(
                (max(1, round(img.width * scale)), max(1, round(img.height * scale))),
                Image.LANCZOS,
            )
        canvas = Image.new("RGB", (img.width + 2 * m, img.height + 2 * m), "white")
        canvas.paste(img, (m, m))
        return canvas
    pw, ph = mm(210), mm(297)  # A4 portrait
    if img.width > img.height:
        pw, ph = ph, pw
    scale = min((pw - 2 * m) / img.width, (ph - 2 * m) / img.height)
    size = (max(1, round(img.width * scale)), max(1, round(img.height * scale)))
    canvas = Image.new("RGB", (pw, ph), "white")
    canvas.paste(img.resize(size, Image.LANCZOS), ((pw - size[0]) // 2, (ph - size[1]) // 2))
    return canvas


def build_pdf(mode: str, files: list[bytes], cfg: dict) -> tuple[bytes, str, int]:
    w = PdfWriter()
    if mode == "img2pdf":
        for b in files:
            buf = io.BytesIO()
            page = to_page(to_rgb(b), cfg["margin"], cfg["page"])
            page.save(buf, format="PDF", resolution=float(PAGE_DPI), quality=90)
            buf.seek(0)
            w.append(buf)
        name = "images.pdf"
    else:
        for b in files:
            w.append(io.BytesIO(b))
        name = "merged.pdf"
    out = io.BytesIO()
    w.write(out)
    return out.getvalue(), name, len(w.pages)


def pdf_info(data: bytes) -> tuple[int, bool]:
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        if doc.needs_pass:
            return 0, True
        return doc.page_count, False
    finally:
        doc.close()


def parse_pages(spec: str, total: int) -> list[int]:
    idx: list[int] = []
    for part in re.split(r"[,\s;]+", spec.strip()):
        if not part:
            continue
        if re.fullmatch(r"\d+", part):
            a = b = int(part)
        else:
            m = re.fullmatch(r"(\d*)[-–](\d*)", part)
            if not m or not (m[1] or m[2]):
                raise UserError("ទម្រង់មិនត្រឹមត្រូវ។ ឧទាហរណ៍៖ 1-3,5,8-10")
            a, b = int(m[1] or 1), int(m[2] or total)
        if a < 1 or b > total or a > b:
            raise UserError(f"លេខទំព័រត្រូវនៅចន្លោះ 1 ដល់ {total}")
        idx.extend(range(a - 1, b))
    if not idx:
        raise UserError("សូមវាយលេខទំព័រ ឧទាហរណ៍៖ 1-3,5,8-10")
    return idx


# Each op: (file_bytes, arg, base_name) -> ([(filename, bytes)], info_text)

def op_word2pdf(data: bytes, ext: str, base: str):
    """បំប្លែង Word → PDF ស្ងាត់ៗពីខាងក្រោយ (Headless 100% មិនលោតផ្ទាំងកម្មវិធី)"""
    soffice = find_soffice()
    suffix = ext if ext in WORD_EXT else ".docx"
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        in_file = tmp_dir / f"input{suffix}"
        out_file = tmp_dir / "input.pdf"
        in_file.write_bytes(data)

        profile_uri = (tmp_dir / "profile").as_uri()
        cmd = [
            soffice,
            f"-env:UserInstallation={profile_uri}",
            "--headless",
            "--invisible",
            "--nocrashreport",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--nologo",
            "--norestore",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(tmp_dir),
            str(in_file),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if res.returncode != 0 or not out_file.exists():
            log.error("LibreOffice failed: %s %s", res.stdout, res.stderr)
            raise UserError("មិនអាចបំប្លែងឯកសារ Word នេះបានទេ (ឯកសារអាចខូច)")

        pdf_bytes = out_file.read_bytes()
        pages, _ = pdf_info(pdf_bytes)
        return (
            [(f"{base}.pdf", pdf_bytes)],
            f"📘 Word → PDF · 📄 {pages} ទំព័រ · 💾 {fmt_size(len(pdf_bytes))}",
        )


def op_pdf2word(data: bytes, _arg, base: str):
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        in_pdf = tmp_dir / "input.pdf"
        out_docx = tmp_dir / "output.docx"
        in_pdf.write_bytes(data)

        doc = pymupdf.open(stream=data, filetype="pdf")
        try:
            total = doc.page_count
            sample = "".join(p.get_text("text") for p in doc)
        finally:
            doc.close()

        used_ocr = False
        if khmer_text_broken(sample):
            # អក្សរក្នុង PDF ខូច ឬជាប្រភេទស្កេន → OCR
            ocr_pdf_to_docx(data, out_docx)
            used_ocr = True
        else:
            cv = Converter(str(in_pdf))
            try:
                cv.convert(str(out_docx))
            finally:
                cv.close()

        if not out_docx.exists():
            raise UserError("មិនអាចបំប្លែង PDF នេះទៅជា Word បានទេ")

        try:
            enforce_khmer_docx_fonts(out_docx, KHMER_FONT)
        except Exception:
            log.warning("Could not enforce Khmer font on docx", exc_info=True)

        docx_bytes = out_docx.read_bytes()
        note = (
            "\n⚠️ ប្រើ OCR ព្រោះអក្សរក្នុង PDF មិនអាចទាញបានត្រឹមត្រូវ (សូមពិនិត្យមើលម្តងទៀត)"
            if used_ocr
            else ""
        )
        return (
            [(f"{base}.docx", docx_bytes)],
            f"📝 PDF → Word ({total} ទំព័រ · {fmt_size(len(docx_bytes))}){note}",
        )


def op_pdf2img(data, _arg, base):
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        total = doc.page_count
        n = min(total, MAX_PAGES)
        files = [
            (f"{base}_page_{i + 1}.jpg", doc[i].get_pixmap(dpi=150).tobytes("jpeg", jpg_quality=90))
            for i in range(n)
        ]
    finally:
        doc.close()
    info = f"🖼 {n} ទំព័រ" + (f" (ពីសរុប {total})" if total > n else "")
    return files, info


def op_split(data, spec, base):
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        total = doc.page_count
        if spec == "all":
            files = []
            for i in range(min(total, MAX_PAGES)):
                part = pymupdf.open()
                part.insert_pdf(doc, from_page=i, to_page=i)
                files.append((f"{base}_page_{i + 1}.pdf", part.tobytes(garbage=3, deflate=True)))
                part.close()
            return files, f"✂️ បំបែកជា {len(files)} ឯកសារ"
        idx = parse_pages(spec, total)
        doc.select(idx)
        return [(f"{base}_pages.pdf", doc.tobytes(garbage=3, deflate=True))], f"✂️ {len(idx)} ទំព័រ"
    finally:
        doc.close()


def op_compress(data, level, base):
    dpi, quality = LEVELS[level]
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        try:
            doc.rewrite_images(dpi_threshold=round(dpi * 1.3), dpi_target=dpi, quality=quality)
        except Exception:
            log.warning("rewrite_images unavailable/failed", exc_info=True)
        out = doc.tobytes(garbage=4, deflate=True, clean=True)
    finally:
        doc.close()
    if len(out) >= len(data):
        return [(f"{base}_compressed.pdf", data)], "ℹ️ ឯកសារនេះបានបង្រួមអស់ហើយ មិនអាចបង្រួមបន្ថែមបានទេ"
    pct = round((1 - len(out) / len(data)) * 100)
    return (
        [(f"{base}_compressed.pdf", out)],
        f"🗜 {fmt_size(len(data))} → {fmt_size(len(out))} (-{pct}%)",
    )


def op_rotate(data, angle, base):
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        for page in doc:
            page.set_rotation((page.rotation + int(angle)) % 360)
        return [(f"{base}_rotated.pdf", doc.tobytes(garbage=3, deflate=True))], f"🔄 បង្វិល {angle}°"
    finally:
        doc.close()


def op_pagenum(data, _arg, base):
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        n = doc.page_count
        for i, page in enumerate(doc):
            text = f"{i + 1} / {n}"
            width = pymupdf.get_text_length(text, fontname="helv", fontsize=10)
            visible = page.rect
            pt = pymupdf.Point((visible.width - width) / 2, visible.height - 20)
            page.insert_text(
                pt * page.derotation_matrix,
                text,
                fontname="helv",
                fontsize=10,
                color=(0.25, 0.25, 0.25),
                rotate=page.rotation,
            )
        return [(f"{base}_numbered.pdf", doc.tobytes(garbage=3, deflate=True))], f"🔢 {n} ទំព័រ"
    finally:
        doc.close()


def op_protect(data, password, base):
    if not password:
        raise UserError("សូមវាយលេខសម្ងាត់")
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        out = doc.tobytes(
            encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw=password, owner_pw=password
        )
    finally:
        doc.close()
    return [(f"{base}_protected.pdf", out)], "🔒 បានដាក់លេខសម្ងាត់"


def op_unlock(data, password, base):
    doc = pymupdf.open(stream=data, filetype="pdf")
    try:
        if not doc.authenticate(password or ""):
            raise UserError("លេខសម្ងាត់មិនត្រឹមត្រូវ សូមព្យាយាមម្តងទៀត")
        out = doc.tobytes(encryption=pymupdf.PDF_ENCRYPT_NONE)
    finally:
        doc.close()
    return [(f"{base}_unlocked.pdf", out)], "🔓 បានដោះលេខសម្ងាត់"


OPS = {
    "pdf2img": op_pdf2img,
    "word2pdf": op_word2pdf,
    "pdf2word": op_pdf2word,
    "split": op_split,
    "compress": op_compress,
    "rotate": op_rotate,
    "pagenum": op_pagenum,
    "protect": op_protect,
    "unlock": op_unlock,
}


# ---------- sending results ----------

async def send_files(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE, files, info: str):
    bot = ctx.bot
    caption = "✅ <b>រួចរាល់!</b>" + (f"\n{info}" if info else "")
    if len(files) == 1:
        name, data = files[0]
        await bot.send_document(
            chat_id, io.BytesIO(data), filename=name, caption=caption, parse_mode=HTML
        )
        return
    for start in range(0, len(files), 10):
        chunk = files[start : start + 10]
        await bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
        if len(chunk) == 1:
            name, data = chunk[0]
            await bot.send_document(
                chat_id,
                io.BytesIO(data),
                filename=name,
                caption=caption if start == 0 else None,
                parse_mode=HTML,
            )
        else:
            media = [
                InputMediaDocument(
                    io.BytesIO(data),
                    filename=name,
                    caption=caption if (start == 0 and i == 0) else None,
                    parse_mode=HTML,
                )
                for i, (name, data) in enumerate(chunk)
            ]
            await bot.send_media_group(chat_id, media)


async def send_menu(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE):
    await ctx.bot.send_message(
        chat_id, "ចង់បន្តធ្វើអ្វីទៀតទេ? 👇", reply_markup=main_menu()
    )


async def run_op(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE, op: str, arg=None):
    s = state(ctx)
    bot = ctx.bot
    data, base = s["pdf"], s["name"] or "document"
    wait = await bot.send_message(chat_id, "⏳ កំពុងដំណើរការ...")
    await bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
    try:
        files, info = await asyncio.to_thread(OPS[op], data, arg, base)
    except UserError as e:
        await wait.edit_text(f"⚠️ {e}")
        return
    except Exception:
        log.exception("op %s failed", op)
        reset(ctx)
        await wait.edit_text(
            "❌ មានបញ្ហា៖ ឯកសារអាចខូច ឬមិនត្រូវទម្រង់។ សូមព្យាយាមម្តងទៀត",
            reply_markup=main_menu(),
        )
        return
    await send_files(chat_id, ctx, files, info)
    await wait.delete()
    reset(ctx)
    await send_menu(chat_id, ctx)


# ---------- status message for multi-file jobs (debounced) ----------

async def refresh_status(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    async with s["lock"]:
        if s["mode"] not in MULTI or not s["files"]:
            return
        if s["status_id"]:
            try:
                await ctx.bot.delete_message(chat_id, s["status_id"])
            except TelegramError:
                pass
            s["status_id"] = None
        n = len(s["files"])
        m = await ctx.bot.send_message(
            chat_id,
            collect_text(s["mode"], n),
            reply_markup=collect_kb(n),
            parse_mode=HTML,
        )
        s["status_id"] = m.message_id


async def _later(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE, my_seq: int, auto: bool):
    await asyncio.sleep(AUTO_DELAY if auto else 1.0)
    if state(ctx)["seq"] != my_seq:
        return
    try:
        if auto:
            await run_done(chat_id, ctx, silent=True)
        else:
            await refresh_status(chat_id, ctx)
    except Exception:
        log.exception("delayed action failed")


def schedule(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE, auto: bool):
    s = state(ctx)
    s["seq"] += 1
    t = asyncio.create_task(_later(chat_id, ctx, s["seq"], auto))
    TASKS.add(t)
    t.add_done_callback(TASKS.discard)


async def run_done(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE, silent: bool = False):
    s = state(ctx)
    bot = ctx.bot
    async with s["lock"]:
        s["seq"] += 1
        if not s["files"] or s["mode"] not in MULTI:
            if not silent:
                await bot.send_message(
                    chat_id,
                    "មិនទាន់មានឯកសារទេ 🤔 សូមជ្រើសរើសឧបករណ៍ 👇",
                    reply_markup=main_menu(),
                )
            return

        mode, files = s["mode"], s["files"]
        s["files"] = []

        if s["status_id"]:
            try:
                await bot.delete_message(chat_id, s["status_id"])
            except TelegramError:
                pass
            s["status_id"] = None

        wait = await bot.send_message(
            chat_id, f"⏳ កំពុងបង្កើត PDF ពី {len(files)} ឯកសារ..."
        )
        await bot.send_chat_action(chat_id, ChatAction.UPLOAD_DOCUMENT)
        try:
            out, name, pages = await asyncio.to_thread(build_pdf, mode, files, dict(cfg_of(ctx)))
        except Exception:
            log.exception("build_pdf failed")
            s["mode"] = None
            await wait.edit_text(
                "❌ មានបញ្ហា៖ ឯកសារណាមួយអាចខូច ឬមិនត្រូវទម្រង់ (ឬ PDF មានលេខសម្ងាត់)។ សូមព្យាយាមម្តងទៀត",
                reply_markup=main_menu(),
            )
            return

        await send_files(
            chat_id, ctx, [(name, out)], f"📄 {pages} ទំព័រ · 💾 {fmt_size(len(out))}"
        )
        await wait.delete()
        if not s["files"]:
            s["mode"] = None
        await send_menu(chat_id, ctx)


# ---------- flow after a single PDF is received ----------

async def after_pdf(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    bot = ctx.bot
    mode = s["mode"]
    pages, encrypted = s["info"]

    if mode == "unlock" and not encrypted:
        reset(ctx)
        await bot.send_message(
            chat_id, "ℹ️ PDF នេះគ្មានលេខសម្ងាត់ទេ", reply_markup=main_menu()
        )
        return
    if mode != "unlock" and encrypted:
        reset(ctx)
        await bot.send_message(
            chat_id,
            "🔒 PDF នេះមានលេខសម្ងាត់ សូមប្រើ «🔓 ដោះលេខសម្ងាត់» សិន",
            reply_markup=main_menu(),
        )
        return

    if mode in ("pdf2img", "pdf2word", "pagenum"):
        await run_op(chat_id, ctx, mode)
    elif mode == "compress":
        kb = Markup(
            [
                [Btn("🟢 គុណភាពខ្ពស់ (បង្រួមតិច)", callback_data="comp:high")],
                [Btn("🟡 មធ្យម (ណែនាំ)", callback_data="comp:mid")],
                [Btn("🔴 បង្រួមខ្លាំង", callback_data="comp:low")],
                [Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")],
            ]
        )
        await bot.send_message(
            chat_id, "🗜 <b>បង្រួម PDF</b>\n\nជ្រើសរើសកម្រិតបង្រួម 👇", reply_markup=kb, parse_mode=HTML
        )
    elif mode == "rotate":
        kb = Markup(
            [
                [
                    Btn("↻ 90° ស្តាំ", callback_data="rot:90"),
                    Btn("🔃 180°", callback_data="rot:180"),
                    Btn("↺ 90° ឆ្វេង", callback_data="rot:270"),
                ],
                [Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")],
            ]
        )
        await bot.send_message(
            chat_id, "🔄 <b>បង្វិល PDF</b>\n\nជ្រើសរើសទិសដៅ 👇", reply_markup=kb, parse_mode=HTML
        )
    elif mode == "split":
        s["await"] = "pages"
        kb = Markup(
            [
                [Btn("📑 បំបែកគ្រប់ទំព័រ", callback_data="split_all")],
                [Btn("⬅️ ត្រឡប់ក្រោយ", callback_data="menu")],
            ]
        )
        await bot.send_message(
            chat_id,
            f"✂️ <b>បំបែក PDF</b>\n\nឯកសារនេះមាន <b>{pages}</b> ទំព័រ\n"
            "វាយលេខទំព័រដែលចង់ទាញយក ឧទាហរណ៍៖\n<code>1-3,5,8-10</code>\n"
            "ឬចុចប៊ូតុងខាងក្រោមដើម្បីបំបែកគ្រប់ទំព័រ",
            reply_markup=kb,
            parse_mode=HTML,
        )
    elif mode in ("protect", "unlock"):
        s["await"] = "password"
        text = (
            "🔒 វាយ <b>លេខសម្ងាត់</b> ដែលចង់ដាក់ 🔑\n(សាររបស់អ្នកនឹងត្រូវលុបវិញដើម្បីសុវត្ថិភាព)"
            if mode == "protect"
            else "🔓 វាយ <b>លេខសម្ងាត់</b> បច្ចុប្បន្នរបស់ PDF នេះ 🔑\n(សាររបស់អ្នកនឹងត្រូវលុបវិញដើម្បីសុវត្ថិភាព)"
        )
        await bot.send_message(chat_id, text, reply_markup=back_kb(), parse_mode=HTML)


async def pdf_choice(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    pages, encrypted = s["info"]
    size = fmt_size(len(s["pdf"]))
    head = (
        f"🔒 <b>ទទួលបាន PDF</b> (មានលេខសម្ងាត់ · {size})"
        if encrypted
        else f"📄 <b>ទទួលបាន PDF</b> ({pages} ទំព័រ · {size})"
    )
    await ctx.bot.send_message(
        chat_id,
        f"{head}\nចង់ធ្វើអ្វីជាមួយវា? 👇",
        reply_markup=pdf_choice_kb(encrypted),
        parse_mode=HTML,
    )


# ---------- handlers ----------

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    reset(ctx)
    await update.message.reply_text(WELCOME, reply_markup=main_menu(), parse_mode=HTML)


async def help_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP, reply_markup=back_kb(), parse_mode=HTML)


async def settings_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text, kb = settings_view(cfg_of(ctx))
    await update.message.reply_text(text, reply_markup=kb, parse_mode=HTML)


async def done_cmd(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await run_done(update.effective_chat.id, ctx)


async def restart_hint(chat_id: int, ctx: ContextTypes.DEFAULT_TYPE):
    await ctx.bot.send_message(
        chat_id, "⚠️ សូមជ្រើសរើសឧបករណ៍ ហើយផ្ញើឯកសារម្តងទៀត 👇", reply_markup=main_menu()
    )


async def on_button(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    chat_id = q.message.chat_id
    s = state(ctx)
    data = q.data

    if data == "menu":
        reset(ctx)
        await safe_edit(q, WELCOME, main_menu())
    elif data == "help":
        await safe_edit(q, HELP, back_kb())
    elif data == "settings":
        text, kb = settings_view(cfg_of(ctx))
        await safe_edit(q, text, kb)
    elif data.startswith("set:"):
        _, key, val = data.split(":")
        cfg = cfg_of(ctx)
        cfg[key] = int(val) if key == "margin" else val
        text, kb = settings_view(cfg)
        await safe_edit(q, text, kb)
    elif data.startswith("mode:"):
        reset(ctx)
        s = state(ctx)
        s["mode"] = data.split(":", 1)[1]
        s["status_id"] = q.message.message_id
        await safe_edit(q, intro_text(s["mode"]), back_kb())
    elif data.startswith("use:"):
        mode = data.split(":", 1)[1]
        if not s["pdf"]:
            await restart_hint(chat_id, ctx)
            return
        try:
            await q.message.delete()
        except TelegramError:
            pass
        s["mode"] = mode
        if mode == "merge":
            s["files"], s["pdf"] = [s["pdf"]], None
            await refresh_status(chat_id, ctx)
        else:
            await after_pdf(chat_id, ctx)
    elif data == "clear":
        s["files"] = []
        s["seq"] += 1
        await safe_edit(q, collect_text(s["mode"] or "img2pdf", 0), back_kb())
    elif data == "done":
        await run_done(chat_id, ctx)
    elif data.startswith(("comp:", "rot:")) or data == "split_all":
        if not s["pdf"]:
            await restart_hint(chat_id, ctx)
            return
        try:
            await q.message.delete()
        except TelegramError:
            pass
        if data == "split_all":
            await run_op(chat_id, ctx, "split", "all")
        elif data.startswith("comp:"):
            await run_op(chat_id, ctx, "compress", data.split(":")[1])
        else:
            await run_op(chat_id, ctx, "rotate", data.split(":")[1])


async def on_text(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    msg = update.message
    chat_id = msg.chat_id
    if s["await"] and s["pdf"]:
        text = (msg.text or "").strip()
        if s["await"] == "password":
            try:
                await msg.delete()
            except TelegramError:
                pass
            await run_op(chat_id, ctx, s["mode"], text)
        else:
            await run_op(chat_id, ctx, "split", text)
        return
    await msg.reply_text("សូមជ្រើសរើសឧបករណ៍ពីម៉ឺនុយខាងក្រោម 👇", reply_markup=main_menu())


def classify(msg) -> tuple[str | None, object, str | None]:
    """Return (kind, telegram_file_holder, original_filename)."""
    if msg.photo:
        return "image", msg.photo[-1], None
    doc = msg.document
    mime = (doc.mime_type or "").lower()
    ext = Path(doc.file_name or "").suffix.lower()
    if mime == "application/pdf" or ext == ".pdf":
        return "pdf", doc, doc.file_name
    if mime in WORD_MIMES or ext in WORD_EXT:
        return "word", doc, doc.file_name
    if mime.startswith("image/") or ext in IMAGE_EXT:
        return "image", doc, doc.file_name
    return None, doc, doc.file_name


async def on_file(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    s = state(ctx)
    msg = update.message
    chat_id = msg.chat_id

    kind, holder, fname = classify(msg)
    if kind is None:
        await msg.reply_text("⚠️ គាំទ្រតែរូបភាព, Word (.docx/.doc) និង PDF ប៉ុណ្ណោះ")
        return
    if getattr(holder, "file_size", None) and holder.file_size > MAX_SIZE:
        await msg.reply_text(f"⚠️ ឯកសារធំពេក (អតិបរមា {MAX_SIZE // (1024 * 1024)}MB)")
        return

    if s["mode"] is None:
        if kind == "image":
            s["mode"] = "img2pdf"
        elif kind == "word":
            s["mode"] = "word2pdf"
        elif msg.media_group_id is not None:
            s["mode"] = "merge"

    mode = s["mode"]
    if mode is not None:
        if mode == "img2pdf":
            wanted, what = "image", "រូបភាព"
        elif mode == "word2pdf":
            wanted, what = "word", "ឯកសារ Word (.docx / .doc)"
        else:
            wanted, what = "pdf", "ឯកសារ PDF"
        if kind != wanted:
            await msg.reply_text(f"⚠️ ឧបករណ៍នេះត្រូវការ{what}ប៉ុណ្ណោះ")
            return

    tg_file = await holder.get_file()
    data = bytes(await tg_file.download_as_bytearray())
    clean_stem = re.sub(r'[\\/:*?"<>|\r\n]+', "_", Path(fname or "document").stem)[:60] or "document"

    if mode == "word2pdf":
        ext = Path(fname or ".docx").suffix.lower() or ".docx"
        s["pdf"] = data
        s["name"] = clean_stem
        await run_op(chat_id, ctx, "word2pdf", ext)
        return

    if mode in MULTI:
        if len(s["files"]) >= MAX_FILES:
            await msg.reply_text(f"⚠️ អតិបរមា {MAX_FILES} ឯកសារ សូមចុច «បង្កើត PDF»")
            return
        if mode == "merge":
            try:
                _, encrypted = await asyncio.to_thread(pdf_info, data)
            except Exception:
                await msg.reply_text("⚠️ ឯកសារ PDF នេះខូច មិនអាចបញ្ចូលគ្នាបានទេ")
                return
            if encrypted:
                await msg.reply_text("🔒 ឯកសារ PDF នេះមានលេខសម្ងាត់ សូមដោះលេខសម្ងាត់ជាមុនសិន")
                return
        async with s["lock"]:
            s["files"].append(data)
        schedule(chat_id, ctx, auto=msg.media_group_id is not None)
        return

    try:
        s["info"] = await asyncio.to_thread(pdf_info, data)
    except Exception:
        log.exception("cannot read pdf")
        await msg.reply_text("⚠️ មិនអាចអានឯកសារ PDF នេះបានទេ (ឯកសារអាចខូច)")
        return
    s["pdf"] = data
    s["name"] = clean_stem

    if mode is None:
        await pdf_choice(chat_id, ctx)
    else:
        await after_pdf(chat_id, ctx)


async def on_error(update: object, ctx: ContextTypes.DEFAULT_TYPE):
    log.error("Unhandled error", exc_info=ctx.error)


async def post_init(app: Application):
    await app.bot.set_my_commands(
        [
            BotCommand("start", "ចាប់ផ្តើម / ម៉ឺនុយ"),
            BotCommand("done", "បញ្ចប់ និងបង្កើត PDF"),
            BotCommand("settings", "ការកំណត់ PDF"),
            BotCommand("help", "ជំនួយ"),
        ]
    )


def main():
    app = (
        Application.builder()
        .token(os.environ["BOT_TOKEN"])
        .post_init(post_init)
        .build()
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(CommandHandler("done", done_cmd))
    app.add_handler(CommandHandler("settings", settings_cmd))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.ALL, on_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_error_handler(on_error)
    log.info("Bot started ✅")
    app.run_polling()


if __name__ == "__main__":
    main()