#!/usr/bin/env python3
"""
بوت تليجرام للبحث في نتائج الثانوية العامة المصرية
Telegram bot for searching Egyptian Thanawya Amma results
"""

import sys
import os
import subprocess
import importlib
import logging

# Auto install dependencies
REQUIRED_PACKAGES = ["openpyxl", "requests", "python-telegram-bot"]
for pkg in REQUIRED_PACKAGES:
    try:
        importlib.import_module(pkg.replace("-", "_"))
    except ImportError:
        print(f"[*] Installing {pkg}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
        except:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q", "--break-system-packages"])
        print(f"[+] {pkg} installed")

import zipfile
import xml.etree.ElementTree as ET
import re
import tempfile
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# ===== CONFIG =====
TOKEN = "8914928931:AAGlH8RPbzKX9MV2nYO37ioSlgd5ldzrEns"
DB_URL = "https://raw.githubusercontent.com/Hhvkvvkv/thanawya-results-db/main/%D9%86%D8%AA%D9%8A%D8%AC%D8%A9%20%D8%AB%D8%A7%D9%86%D9%88%D9%8A%D8%A9%20%D8%B9%D8%A7%D9%85%D8%A9%20%D9%86%D8%B8%D8%A7%D9%85%20%D8%AD%D8%AF%D9%8A%D8%AB.xlsx"

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ===== DATABASE =====
class ThanawyaDB:
    def __init__(self):
        self.shared_strings = []
        self.loaded = False
        self.db_path = None

    def load(self):
        if self.loaded:
            return True

        # Try to find existing DB in current dir, temp dir, or download
        possible_paths = [
            os.path.join(os.getcwd(), "نتيجة ثانوية عامة نظام حديث.xlsx"),
            os.path.join(tempfile.gettempdir(), "نتيجة ثانوية عامة نظام حديث.xlsx"),
        ]

        for p in possible_paths:
            if os.path.exists(p):
                self.db_path = p
                break

        if not self.db_path:
            import requests
            self.db_path = possible_paths[1]  # save to temp
            print("[*] Downloading database from GitHub (40MB)...")
            r = requests.get(DB_URL, stream=True, timeout=300)
            r.raise_for_status()
            total = 0
            with open(self.db_path, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        total += len(chunk)
                        if total % 5000000 == 0:
                            print(f"  Downloaded {total//1048576} MB...")
            print(f"[+] Downloaded to {self.db_path}")

        print("[*] Loading database...")
        z = zipfile.ZipFile(self.db_path)
        ss_xml = z.read('xl/sharedStrings.xml')
        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        tree = ET.fromstring(ss_xml)
        for si in tree.findall(f'.//{ns}si'):
            text_parts = []
            for t in si.findall(f'.//{ns}t'):
                if t.text:
                    text_parts.append(t.text)
            self.shared_strings.append(''.join(text_parts))
        z.close()
        self.loaded = True
        print(f"[+] Loaded {len(self.shared_strings)} student names")
        return True

    def search(self, name_query):
        query_parts = name_query.strip().split()
        results = []
        for idx, name in enumerate(self.shared_strings):
            if all(part in name for part in query_parts):
                results.append((idx, name))
        return results[:15]

    def get_student_result(self, ss_index):
        z = zipfile.ZipFile(self.db_path)
        raw = z.read('xl/worksheets/sheet1.xml')
        z.close()
        pos = raw.find(f'<v>{ss_index}</v>'.encode())
        if pos == -1:
            return None
        rs = raw.rfind(b'<row', 0, pos)
        re_ = raw.find(b'</row>', pos) + 6
        row = raw[rs:re_].decode('utf-8', errors='replace')
        result = {}
        m = re.search(r'row r="(\d+)"', row)
        result['row'] = m.group(1) if m else '?'
        m = re.search(r'<c r="A\d+"><v>([^<]+)</v>', row)
        if m:
            result['seating'] = m.group(1)
        result['name'] = self.shared_strings[ss_index] if ss_index < len(self.shared_strings) else str(ss_index)
        m = re.search(r'<c r="C\d+"><v>([^<]+)</v>', row)
        if m:
            result['total'] = float(m.group(2))
        m = re.search(r'<c r="D\d+" t="s"><v>(\d+)</v>', row)
        if m:
            ci = int(m.group(2))
            result['status'] = self.shared_strings[ci] if ci < len(self.shared_strings) else '?'
        return result

    def search_with_results(self, query):
        if not self.loaded:
            self.load()
        matches = self.search(query)
        results = []
        for idx, name in matches:
            r = self.get_student_result(idx)
            if r:
                results.append(r)
        return results

db = ThanawyaDB()

# ===== BOT HANDLERS =====
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🎓 *بوت البحث في نتائج الثانوية العامة*\n\n"
        "أرسل اسم الطالب للبحث عن نتيجته\n"
        "مثال: `ناديه محمد عبد المنعم`\n\n"
        "✅ قاعدة البيانات محملة – جاهز للبحث",
        parse_mode="Markdown"
    )

async def search_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.message.text.strip()
    if not query:
        await update.message.reply_text("❌ من فضلك أدخل اسم الطالب")
        return

    msg = await update.message.reply_text(
        f"🔍 جاري البحث عن: `{query}`\n⏳ انتظر قليلاً...",
        parse_mode="Markdown"
    )

    try:
        results = db.search_with_results(query)
    except Exception as e:
        await msg.edit_text(f"❌ خطأ: {e}\n\nحاول مرة أخرى")
        logger.exception("Search error")
        return

    if not results:
        await msg.edit_text(
            f"❌ لا توجد نتائج للاسم: `{query}`\n\nحاول بصيغة مختلفة",
            parse_mode="Markdown"
        )
        return

    text = f"✅ *نتائج البحث عن:* `{query}`\n\n"
    for i, r in enumerate(results, 1):
        name = r.get('name', '?')
        seating = r.get('seating', '?')
        total = r.get('total', '?')
        status = r.get('status', '')
        pct = f" - نسبته: `{(total/320)*100:.2f}%`" if isinstance(total, (int, float)) else ""
        status_text = f"\n   📌 الحالة: {status}" if status else ""
        text += f"*{i}.* {name}\n   🆔 {seating} - المجموع: {total}{pct}{status_text}\n\n"

    text += "🔹 أرسل اسمًا جديدًا للبحث"

    if len(text) > 4000:
        text = text[:3500] + "\n\n✂️ ... نتائج مختصرة (أكثر من 4000 حرف)"

    await msg.edit_text(text, parse_mode="Markdown")

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    logger.warning(f"Update {update} caused error {context.error}")

def main():
    print("[*] Starting Thanawya Results Telegram Bot...")
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, search_handler))
    app.add_error_handler(error_handler)
    print("[+] Bot is running! Send a message on Telegram...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
