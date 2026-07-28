#!/usr/bin/env python3
"""
بوت تليجرام للبحث في نتائج الثانوية العامة المصرية
يدعم Kaggle / Colab / أي سيرفر
"""

import sys
import os
import subprocess
import importlib
import logging

# ========== AUTO INSTALL ==========
REQUIRED_PACKAGES = ["openpyxl", "requests", "python-telegram-bot"]
for pkg in REQUIRED_PACKAGES:
    module_name = pkg.replace("-", "_").replace("python_", "")
    try:
        importlib.import_module(module_name)
    except ImportError:
        print(f"[*] Installing {pkg}...")
        for flag in [["-q"], ["-q", "--break-system-packages"], ["-q", "--user"]]:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg] + flag)
                print(f"[+] {pkg} installed")
                break
            except:
                continue

import zipfile
import xml.etree.ElementTree as ET
import re
import tempfile
import traceback
import asyncio

from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# ========== CONFIG ==========
BOT_TOKEN = "8914928931:AAGlH8RPbzKX9MV2nYO37ioSlgd5ldzrEns"
DB_FILENAME = "نتيجة ثانوية عامة نظام حديث.xlsx"
# Use 'master' not 'main'
DB_URL = "https://raw.githubusercontent.com/Hhvkvvkv/thanawya-results-db/master/%D9%86%D8%AA%D9%8A%D8%AC%D8%A9%20%D8%AB%D8%A7%D9%86%D9%88%D9%8A%D8%A9%20%D8%B9%D8%A7%D9%85%D8%A9%20%D9%86%D8%B8%D8%A7%D9%85%20%D8%AD%D8%AF%D9%8A%D8%AB.xlsx"

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ========== DB ==========
class ResultDB:
    def __init__(self):
        self.names = []
        self.path = None
        self.ready = False

    def find_or_download_db(self):
        search = [
            os.path.join(os.getcwd(), DB_FILENAME),
            os.path.join(tempfile.gettempdir(), DB_FILENAME),
            "/tmp/" + DB_FILENAME,
            "/kaggle/working/" + DB_FILENAME,
            "/content/" + DB_FILENAME,
        ]
        for p in search:
            if os.path.exists(p):
                self.path = p
                return True
        # Download
        try:
            import requests
            self.path = search[1]
            print("[*] Downloading DB (~40MB)...", flush=True)
            r = requests.get(DB_URL, stream=True, timeout=600)
            r.raise_for_status()
            with open(self.path, "wb") as f:
                for c in r.iter_content(8192):
                    if c:
                        f.write(c)
            print("[+] Downloaded DB", flush=True)
            return True
        except Exception as e:
            print(f"[-] Download error: {e}", flush=True)
            return False

    def load_names(self):
        if self.ready:
            return True
        if not self.path or not os.path.exists(self.path):
            if not self.find_or_download_db():
                return False
        try:
            z = zipfile.ZipFile(self.path)
            xml = z.read("xl/sharedStrings.xml")
            z.close()
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            root = ET.fromstring(xml)
            for si in root.findall(f".//{ns}si"):
                parts = [t.text or "" for t in si.findall(f".//{ns}t")]
                self.names.append("".join(parts))
            self.ready = True
            print(f"[+] Loaded {len(self.names)} names", flush=True)
            return True
        except Exception as e:
            print(f"[-] Load error: {e}", flush=True)
            return False

    def search(self, q):
        parts = q.strip().split()
        out = []
        for i, n in enumerate(self.names):
            if all(p in n for p in parts):
                out.append((i, n))
                if len(out) >= 15:
                    break
        return out

    def get_result(self, idx):
        try:
            z = zipfile.ZipFile(self.path)
            sheet = z.read("xl/worksheets/sheet1.xml")
            z.close()
            pos = sheet.find(f"<v>{idx}</v>".encode())
            if pos < 0:
                return None
            rs = sheet.rfind(b"<row", 0, pos)
            re2 = sheet.find(b"</row>", pos) + 6
            xml = sheet[rs:re2].decode("utf-8", errors="replace")
            out = {}
            m = re.search(r'row r="(\d+)"', xml)
            if m: out["row"] = m.group(1)
            m = re.search(r'<c r="A\d+"><v>([^<]+)</v>', xml)
            if m: out["seating"] = m.group(1)
            m = re.search(r'<c r="C\d+"><v>([^<]+)</v>', xml)
            if m:
                try: out["total"] = float(m.group(1))
                except: out["total"] = m.group(1)
            m = re.search(r'<c r="D\d+" t="s"><v>(\d+)</v>', xml)
            if m:
                si = int(m.group(1))
                out["status"] = self.names[si] if si < len(self.names) else ""
            out["name"] = self.names[idx] if idx < len(self.names) else "?"
            return out
        except:
            return None

    def search_results(self, q):
        if not self.ready and not self.load_names():
            return None
        matches = self.search(q)
        if not matches:
            return []
        out = []
        for i, _ in matches:
            r = self.get_result(i)
            if r:
                out.append(r)
        return out

db = ResultDB()

# ========== HANDLERS ==========
async def start(upd: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await upd.message.reply_text(
        "🎓 *بوت نتائج الثانوية العامة*\n\n"
        "أرسل اسم الطالب للبحث\n"
        "مثال: `ناديه محمد عبد المنعم`",
        parse_mode="Markdown"
    )

async def search(upd: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = upd.message.text.strip()
    if not q:
        await upd.message.reply_text("❌ أدخل اسمًا")
        return
    await upd.message.chat.send_action(action="typing")
    try:
        res = db.search_results(q)
    except Exception as e:
        await upd.message.reply_text(f"❌ خطأ: {e}")
        return
    if res is None:
        await upd.message.reply_text("❌ فشل تحميل قاعدة البيانات")
        return
    if not res:
        await upd.message.reply_text(
            f"❌ لا توجد نتائج لـ: `{q}`\n💡 جرب اسمًا آخر",
            parse_mode="Markdown"
        )
        return
    text = f"✅ نتائج: `{q}`\n\n"
    for i, r in enumerate(res, 1):
        n = r.get("name", "?")
        s = r.get("seating", "?")
        t = r.get("total", "?")
        st = r.get("status", "")
        text += f"{i}. {n}\n   🆔 {s}"
        if isinstance(t, (int, float)):
            text += f" - {t}/320 ({t/320*100:.1f}%)"
        elif t:
            text += f" - {t}"
        if st:
            text += f" | {st}"
        text += "\n\n"
    if len(text) > 4000:
        text = text[:3500] + "\n...(مختصر)"
    await upd.message.reply_text(text)

async def help_cmd(upd: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await upd.message.reply_text("📖 أرسل اسم الطالب للبحث عن نتيجته")

async def err(upd: object, ctx: ContextTypes.DEFAULT_TYPE):
    logger.error(f"Error: {ctx.error}")

# ========== MAIN ==========
def main():
    print("Starting Thanawya Bot...", flush=True)
    db.load_names()

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, search))
    app.add_error_handler(err)

    print("Bot running! Press Ctrl+C to stop.", flush=True)

    # Fix for asyncio.run() in running event loop (Kaggle/Colab/Jupyter)
    try:
        loop = asyncio.get_running_loop()
        # We're in a running loop (Kaggle/Colab)
        print("Running in existing event loop...", flush=True)
        loop.create_task(app.run_polling(allowed_updates=Update.ALL_TYPES))
    except RuntimeError:
        # No running loop - normal case
        app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
