#!/usr/bin/env python3
"""
بوت البحث في نتائج الثانوية العامة المصرية
Thanawya Amma Results Search Bot
"""

import sys
import os
import subprocess
import importlib

# Auto install dependencies
REQUIRED_PACKAGES = ["openpyxl", "requests"]
for pkg in REQUIRED_PACKAGES:
    try:
        importlib.import_module(pkg)
    except ImportError:
        print(f"[*] Installing {pkg}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q", "--break-system-packages"])
        print(f"[+] {pkg} installed")

import openpyxl
import zipfile
import xml.etree.ElementTree as ET
import re
import json

# ===== CONFIG =====
# GitHub raw URL for the database
DB_URL = "https://raw.githubusercontent.com/Hhvkvvkv/thanawya-results-db/main/%D9%86%D8%AA%D9%8A%D8%AC%D8%A9%20%D8%AB%D8%A7%D9%86%D9%88%D9%8A%D8%A9%20%D8%B9%D8%A7%D9%85%D8%A9%20%D9%86%D8%B8%D8%A7%D9%85%20%D8%AD%D8%AF%D9%8A%D8%AB.xlsx"
LOCAL_DB = os.path.join(os.path.dirname(__file__), "نتيجة ثانوية عامة نظام حديث.xlsx")

# ===== DATABASE LOADER =====
class ThanawyaDB:
    def __init__(self, path=None):
        self.path = path or LOCAL_DB
        self.shared_strings = []
        self.loaded = False

    def load(self):
        """Load shared strings index from the xlsx file"""
        if self.loaded:
            return True
        
        if not os.path.exists(self.path):
            print(f"[-] Database not found at: {self.path}")
            print("[*] Downloading from GitHub...")
            try:
                import requests
                r = requests.get(DB_URL, stream=True, timeout=60)
                r.raise_for_status()
                with open(self.path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
                print(f"[+] Downloaded to {self.path}")
            except Exception as e:
                print(f"[-] Failed to download: {e}")
                return False

        print("[*] Loading shared strings (may take a moment)...")
        try:
            z = zipfile.ZipFile(self.path)
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
        except Exception as e:
            print(f"[-] Error loading database: {e}")
            return False

    def search(self, name_query, exact=False):
        """Search for a student by name"""
        if not self.loaded:
            if not self.load():
                return []

        results = []
        query_parts = name_query.strip().split()
        
        for idx, name in enumerate(self.shared_strings):
            if exact:
                if name == name_query:
                    results.append((idx, name))
            else:
                if all(part in name for part in query_parts):
                    results.append((idx, name))
        
        return results

    def get_student_result(self, ss_index):
        """Get the full result for a student by their shared string index"""
        try:
            z = zipfile.ZipFile(self.path)
            raw = z.read('xl/worksheets/sheet1.xml')
            z.close()

            # Search for this SS index in column B
            search_bytes = f'<v>{ss_index}</v>'.encode()
            pos = raw.find(search_bytes)
            if pos == -1:
                return None

            # Find the enclosing row
            row_start = raw.rfind(b'<row', 0, pos)
            row_end = raw.find(b'</row>', pos) + 6
            row_xml = raw[row_start:row_end].decode('utf-8', errors='replace')

            # Parse row
            result = {}
            rm = re.search(r'row r="(\d+)"', row_xml)
            result['row'] = rm.group(1) if rm else '?'

            # Column A - seating number
            m = re.search(r'<c r="A\d+"><v>([^<]+)</v>', row_xml)
            if m:
                result['seating_no'] = m.group(1)

            # Column B - name (already have it from SS index)
            result['name'] = self.shared_strings[ss_index] if ss_index < len(self.shared_strings) else f'Index {ss_index}'

            # Column C - total degree
            m = re.search(r'<c r="C(\d+)"><v>([^<]+)</v>', row_xml)
            if m:
                result['total_degree'] = float(m.group(2))

            # Column D - case (ناجح/راسب)
            m = re.search(r'<c r="D(\d+)" t="s"><v>(\d+)</v>', row_xml)
            if m:
                case_idx = int(m.group(2))
                result['status'] = self.shared_strings[case_idx] if case_idx < len(self.shared_strings) else f'Index {case_idx}'

            return result

        except Exception as e:
            print(f"[-] Error getting result: {e}")
            return None

    def search_with_results(self, query, max_results=10):
        """Search students and get their full results"""
        matches = self.search(query)
        if not matches:
            return []
        
        results = []
        for idx, name in matches[:max_results]:
            result = self.get_student_result(idx)
            if result:
                results.append(result)
        return results


# ===== MAIN =====
def main():
    db = ThanawyaDB()
    
    if len(sys.argv) > 1:
        query = ' '.join(sys.argv[1:])
        print(f"\n🔍 Searching for: {query}")
        print("=" * 50)
        
        results = db.search_with_results(query)
        
        if not results:
            print("❌ No results found. Try a different spelling.")
            return
        
        for i, r in enumerate(results, 1):
            print(f"\n{'─' * 40}")
            print(f"📌 Result #{i}")
            print(f"{'─' * 40}")
            print(f"  الاسم:       {r.get('name', '?')}")
            print(f"  رقم الجلوس:  {r.get('seating_no', '?')}")
            print(f"  المجموع:     {r.get('total_degree', '?')}")
            if 'status' in r:
                print(f"  الحالة:      {r['status']}")
            
            total = r.get('total_degree')
            if total and isinstance(total, (int, float)):
                pct = (total / 320) * 100
                print(f"  النسبة:      {pct:.2f}%")
        print()
    else:
        print("\n" + "=" * 50)
        print("  🏫 بوت البحث في نتائج الثانوية العامة")
        print("  Thanawya Amma Results Search Bot")
        print("=" * 50)
        print("\n  الاستخدام:")
        print(f"    python3 {os.path.basename(__file__)} \"اسم الطالب\"")
        print("\n  أمثلة:")
        print(f"    python3 {os.path.basename(__file__)} ناديه محمد عبد المنعم")
        print(f"    python3 {os.path.basename(__file__)} محمد رضا")
        print("\n  ✅ البوت سيقوم بتحميل قاعدة البيانات تلقائيا")
        print("     عند أول تشغيل إذا لم تكن موجودة محليا\n")

if __name__ == "__main__":
    main()
