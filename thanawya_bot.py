#!/usr/bin/env python3
"""
بوت تليجرام للبحث في نتائج الثانوية العامة المصرية
يدعم التشغيل على Kaggle / أي سيرفر مع تحميل تلقائي للمكتبات
"""

import sys
import os
import subprocess
import importlib
import logging
import asyncio

# ========== AUTO INSTALL DEPENDENCIES ==========
REQUIRED_PACKAGES = ["openpyxl", "requests", "python-telegram-bot"]
for pkg in REQUIRED_PACKAGES:
    module_name = pkg.replace("-", "_").replace("python_", "")
    try:
        importlib.import_module(module_name)
    except ImportError:
        print(f"[*] Installing {pkg}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
        except:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q", "--break-system-packages"])
            except:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q", "--user"])
        print(f"[+] {pkg} installed")

# ========== IMPORTS ==========
import zipfile
import xml.etree.ElementTree as ET
import re
import tempfile
import json
import traceback

from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# ========== CONFIG ==========
BOT_TOKEN = "8914928931:AAGlH8RPbzKX9MV2nYO37ioSlgd5ldzrEns"
DB_FILENAME = "نتيجة ثانوية عامة نظام حديث.xlsx"
DB_URL = "https://raw.githubusercontent.com/Hhvkvvkv/thanawya-results-db/main/%D9%86%D8%AA%D9%8A%D8%AC%D8%A9%20%D8%AB%D8%A7%D9%86%D9%88%D9%8A%D8%A9%20%D8%B9%D8%A7%D9%85%D8%A9%20%D9%86%D8%B8%D8%A7%D9%85%20%D8%AD%D8%AF%D9%8A%D8%AB.xlsx"

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ========== DATABASE ==========
class ResultDB:
    """Handles loading and searching the Excel database"""

    def __init__(self):
        self.names = []       # list of all student names
        self.path = None      # path to xlsx file
        self.ready = False

    def find_or_download_db(self):
        """Find existing DB or download from GitHub"""
        search_paths = [
            os.path.join(os.getcwd(), DB_FILENAME),
            os.path.join(tempfile.gettempdir(), DB_FILENAME),
            os.path.join("/tmp", DB_FILENAME),
            os.path.join("/content", DB_FILENAME),  # Colab
            os.path.join("/kaggle/working", DB_FILENAME),  # Kaggle
        ]

        for p in search_paths:
            if os.path.exists(p):
                self.path = p
                print(f"[+] Found DB at: {p}")
                return True

        # Download
        print("[*] Downloading database from GitHub (~40MB)...")
        try:
            import requests
            self.path = search_paths[1]  # /tmp
            r = requests.get(DB_URL, stream=True, timeout=300)
            r.raise_for_status()
            with open(self.path, "wb") as f:
                downloaded = 0
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
            print(f"[+] Downloaded to: {self.path}")
            return True
        except Exception as e:
            print(f"[-] Download failed: {e}")
            return False

    def load_names(self):
        """Load all student names from shared strings XML"""
        if self.ready:
            return True

        if not self.path or not os.path.exists(self.path):
            if not self.find_or_download_db():
                return False

        print("[*] Loading student names...")
        try:
            z = zipfile.ZipFile(self.path)
            xml_data = z.read("xl/sharedStrings.xml")
            z.close()

            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            root = ET.fromstring(xml_data)

            for si in root.findall(f".//{ns}si"):
                parts = []
                for t in si.findall(f".//{ns}t"):
                    if t.text:
                        parts.append(t.text)
                self.names.append("".join(parts))

            self.ready = True
            print(f"[+] Loaded {len(self.names)} student names")
            return True

        except Exception as e:
            print(f"[-] Failed to load DB: {e}")
            traceback.print_exc()
            return False

    def search(self, query):
        """Search for students by name parts"""
        parts = query.strip().split()
        matches = []
        for idx, name in enumerate(self.names):
            if all(p in name for p in parts):
                matches.append((idx, name))
                if len(matches) >= 15:
                    break
        return matches

    def get_result(self, name_index):
        """Get full result row for a student by their name index"""
        try:
            z = zipfile.ZipFile(self.path)
            sheet = z.read("xl/worksheets/sheet1.xml")
            z.close()

            # Find the row containing this name index in column B
            search = f'<v>{name_index}</v>'.encode("utf-8")
            pos = sheet.find(search)
            if pos == -1:
                return None

            # Extract the full <row> tag
            row_start = sheet.rfind(b"<row", 0, pos)
            row_end = sheet.find(b"</row>", pos) + 6
            if row_start == -1 or row_end == -1:
                return None

            row_xml = sheet[row_start:row_end].decode("utf-8", errors="replace")

            # Parse row number
            row_num = ""
            m = re.search(r'row r="(\d+)"', row_xml)
            if m:
                row_num = m.group(1)

            # Parse column A (seating number)
            seating = ""
            m = re.search(r'<c r="A\d+"><v>([^<]+)</v>', row_xml)
            if m:
                seating = m.group(1)

            # Parse column C (total degree)
            total = None
            m = re.search(r'<c r="C\d+"><v>([^<]+)</v>', row_xml)
            if m:
                try:
                    total = float(m.group(1))
                except:
                    total = m.group(1)

            # Parse column D (status - shared string index)
            status = ""
            m = re.search(r'<c r="D\d+" t="s"><v>(\d+)</v>', row_xml)
            if m:
                si = int(m.group(1))
                if si < len(self.names):
                    status = self.names[si]

            name = self.names[name_index] if name_index < len(self.names) else "?"

            return {
                "row": row_num,
                "seating": seating,
                "name": name,
                "total": total,
                "status": status,
            }

        except Exception as e:
            logger.exception(f"Error getting result for index {name_index}")
            return None

    def search_with_results(self, query):
        """Full search: find matches and return their results"""
        if not self.ready:
            if not self.load_names():
                return None

        matches = self.search(query)
        if not matches:
            return []

        results = []
        for idx, _ in matches:
            r = self.get_result(idx)
            if r:
                results.append(r)
        return results


# ========== GLOBAL DB INSTANCE ==========
db = ResultDB()


# ========== BOT HANDLERS ==========
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Welcome message"""
    await update.message.reply_text(
        "🎓 *بوت البحث في نتائج الثانوية العامة*\n\n"
        "أرسل اسم الطالب للبحث عن نتيجته\n"
        "مثال: `ناديه محمد عبد المنعم`\n\n"
        "مدعوم من قاعدة بيانات وزارة التربية والتعليم",
        parse_mode="Markdown"
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Search for a student by name"""
    query = update.message.text.strip()
    if not query:
        await update.message.reply_text("❌ من فضلك أدخل اسم الطالب")
        return

    # Send typing action and initial message
    await update.message.chat.send_action(action="typing")
    msg = await update.message.reply_text(
        f"🔍 جاري البحث عن: `{query}`\n⏳ يرجى الانتظار...",
        parse_mode="Markdown"
    )

    try:
        results = db.search_with_results(query)
    except Exception as e:
        logger.exception("Search error")
        await msg.edit_text(f"❌ حدث خطأ:\n`{e}`", parse_mode="Markdown")
        return

    if results is None:
        await msg.edit_text("❌ فشل تحميل قاعدة البيانات. حاول مرة أخرى")
        return

    if not results:
        await msg.edit_text(
            f"❌ لا توجد نتائج للاسم: `{query}`\n\n"
            "💡 حاول بصيغة مختلفة\n"
            "مثال: `محمد رضا`",
            parse_mode="Markdown"
        )
        return

    # Build response
    text = f"✅ *نتائج البحث عن:* `{query}`\n\n"
    for i, r in enumerate(results, 1):
        name = r.get("name", "?")
        seating = r.get("seating", "?")
        total = r.get("total", "?")
        status = r.get("status", "")

        line = f"*{i}.* {name}\n   🆔 {seating}"

        if isinstance(total, (int, float)):
            line += f" - المجموع: {total}"
            pct = (total / 320) * 100
            line += f" ({pct:.2f}%)"
        else:
            if total:
                line += f" - {total}"

        if status:
            line += f"\n   📌 {status}"

        line += "\n\n"
        text += line

    text += "🔹 أرسل اسمًا آخر للبحث"

    if len(text) > 4000:
        text = text[:3900] + "\n\n... (اختصار بسبب طول النص)"

    await msg.edit_text(text, parse_mode="Markdown")


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Help message"""
    await update.message.reply_text(
        "📖 *مساعدة البوت*\n\n"
        "• أرسل *اسم الطالب* للبحث\n"
        "• مثال: `ناديه محمد عبد المنعم`\n"
        "• مثال: `محمد رضا`\n"
        "• مثال: `أحمد السيد`\n\n"
        "البيانات تشمل:\n"
        "• الاسم الرباعي\n"
        "• رقم الجلوس\n"
        "• المجموع والنسبة المئوية\n"
        "• حالة النجاح",
        parse_mode="Markdown"
    )


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE):
    """Log errors"""
    logger.error(f"Exception while handling an update: {context.error}")


# ========== MAIN ==========
def main():
    """Start the bot"""
    print("=" * 50)
    print("  🏫 Thanawya Results Telegram Bot")
    print("=" * 50)

    # Pre-load database
    print("[*] Initializing database...")
    if db.load_names():
        print("[+] Database ready!")
    else:
        print("[-] Database not available. Bot will try on first search.")

    # Build application
    print("[*] Starting bot...")
    app = Application.builder().token(BOT_TOKEN).build()

    # Register handlers
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_error_handler(error_handler)

    print("[+] Bot is running! Press Ctrl+C to stop.")
    print("[+] Open Telegram and send a message to the bot.")

    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES)
    except RuntimeError as e:
        if "event loop" in str(e).lower():
            # Kaggle/Colab workaround: use asyncio directly
            print("[*] Using asyncio workaround for this environment...")
            asyncio.run(_async_main(app))
        else:
            raise


async def _async_main(app):
    """Async entry point for Kaggle/Colab"""
    async with app:
        await app.start()
        print("[+] Bot polling started (async mode)")
        # Keep running
        while True:
            await asyncio.sleep(3600)  # Sleep 1 hour, will be interrupted by Ctrl+C


if __name__ == "__main__":
    main()
