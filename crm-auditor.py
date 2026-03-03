import os
import csv
import re
import difflib
from collections import defaultdict

class CRMAuditor:
    def __init__(self, root_dir="."):
        self.root_dir = root_dir
        self.schema_cols = set()
        self.valid_outcomes = set()
        self.valid_stages = set()
        self.valid_statuses = set()
        self.settings_map = defaultdict(set)
        
        # Files to look for
        self.schema_file = "System_Schema.csv"
        self.settings_file = "Settings.csv"
        
        # Violations Report
        self.violations = []
        self.function_map = defaultdict(list)

    def load_config(self):
        """Loads the Source of Truth (Schema & Settings)"""
        print("🔵 Loading Configuration...")
        
        # 1. Load Schema
        # We look recursively just in case they are in a subfolder
        schema_path = self.find_file(self.schema_file)
        if schema_path:
            with open(schema_path, 'r', encoding='utf-8', errors='replace') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if 'Label' in row and row['Label']:
                        self.schema_cols.add(row['Label'].strip())
                    if 'API_Name' in row and row['API_Name']:
                        self.schema_cols.add(row['API_Name'].strip())
            print(f"   ✅ Loaded {len(self.schema_cols)} Schema Columns from {schema_path}")
        else:
            print(f"   ❌ CRITICAL: {self.schema_file} not found. Audit will be limited.")

        # 2. Load Settings
        settings_path = self.find_file(self.settings_file)
        if settings_path:
            with open(settings_path, 'r', encoding='utf-8', errors='replace') as f:
                reader = csv.reader(f, delimiter='\t')
                next(reader, None) # Skip header
                for row in reader:
                    if len(row) < 2: continue
                    category = row[0].strip()
                    key = row[1].strip()
                    
                    if category == 'VALIDATION_LIST' and key == 'Outcomes':
                        raw_values = row[2:]
                        for val in raw_values:
                            for item in val.split(','):
                                self.valid_outcomes.add(item.strip())
                                
                    elif category == 'VALIDATION_LIST' and key == 'Stages':
                        raw_values = row[2:]
                        for val in raw_values:
                            for item in val.split(','):
                                self.valid_stages.add(item.strip())

                    elif category == 'VALIDATION_LIST' and key == 'Statuses':
                        raw_values = row[2:]
                        for val in raw_values:
                            for item in val.split(','):
                                self.valid_statuses.add(item.strip())
                                
                    self.settings_map[category].add(key)
            
            print(f"   ✅ Loaded Settings from {settings_path}")
        else:
            print(f"   ❌ CRITICAL: {self.settings_file} not found.")

    def find_file(self, filename):
        for dirpath, _, filenames in os.walk(self.root_dir):
            if filename in filenames:
                return os.path.join(dirpath, filename)
        return None

    def scan_codebase(self):
        """Walks through JS, GS, AND HTML files."""
        print("\n🔵 Scanning Codebase (JS, GS, HTML)...")
        
        # Regex to capture string literals: "text" or 'text'
        string_pattern = re.compile(r'["\']([^"\']+)["\']')
        # Regex to capture function names
        func_pattern = re.compile(r'function\s+(\w+)')

        exclude_dirs = {'.git', 'node_modules', '__pycache__', '.idea', 'venv'}

        for dirpath, dirnames, filenames in os.walk(self.root_dir):
            dirnames[:] = [d for d in dirnames if d not in exclude_dirs]
            
            for filename in filenames:
                # UPDATED: Now includes .html
                if filename.lower().endswith(('.js', '.gs', '.html')):
                    filepath = os.path.join(dirpath, filename)
                    self.analyze_file(filepath, string_pattern, func_pattern)

    def analyze_file(self, filepath, string_pattern, func_pattern):
        rel_path = os.path.relpath(filepath, self.root_dir)
        
        try:
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.readlines()
        except Exception as e:
            print(f"   ⚠️ Error reading {rel_path}: {e}")
            return

        for i, line in enumerate(lines):
            line_num = i + 1
            code = line.strip()
            
            # 1. Capture Function Names
            func_match = func_pattern.search(code)
            if func_match:
                self.function_map[rel_path].append(func_match.group(1))

            # 2. Extract Strings to Check Validity
            strings_found = string_pattern.findall(code)
            for s in strings_found:
                self.validate_string_literal(s, rel_path, line_num, code)

    def validate_string_literal(self, text, filename, line_num, code_context):
        """Core Logic: Checks if a string in code violates the Schema or Settings."""
        
        if len(text) < 3: return
        # Common false positives
        if text in ['red', 'blue', 'GET', 'POST', 'application/json', 'text/html', 'center']: return
        
        # --- CHECK 1: "Not Contacted" Usage ---
        # It is strictly banned in Outreach logic or HTML buttons intended for logging
        if text == "Not Contacted":
            # If it's in an HTML file, it's likely a button label or value
            if filename.endswith(".html"):
                 self.add_violation(filename, line_num, "HIGH", 
                                   "Found 'Not Contacted' in HTML UI. This status should be removed/replaced with 'No Answer'.", code_context)
            # If it's in Outreach logic
            elif "Outreach" in filename or "outreach" in code_context.lower():
                self.add_violation(filename, line_num, "HIGH", 
                                   "Usage of 'Not Contacted' in Outreach logic. (Should be 'No Answer'?)", code_context)

        # --- CHECK 2: Schema Column Mismatches (Typos) ---
        for col in self.schema_cols:
            if text != col and self.is_typo(text, col):
                self.add_violation(filename, line_num, "MEDIUM", 
                                   f"Potential Schema Typo: Found '{text}', expected '{col}'", code_context)

        # --- CHECK 3: Invalid Outcome/Stage/Status Logic ---
        # Context clues to know if we are talking about CRM data
        is_crm_context = any(x in code_context for x in ["Outcome", "Stage", "Status", "val", "set", "if"])
        
        if is_crm_context:
            if self.looks_like_category(text, self.valid_outcomes) and text not in self.valid_outcomes:
                 # Reduce noise: don't flag if it matches a variable name like "outcome"
                 if text.lower() != "outcome":
                    self.add_violation(filename, line_num, "HIGH", 
                                   f"Invalid Outcome string: '{text}'. Not in Settings.tsv.", code_context)
            
            elif self.looks_like_category(text, self.valid_stages) and text not in self.valid_stages:
                 if text.lower() != "stage":
                    self.add_violation(filename, line_num, "HIGH", 
                                   f"Invalid Stage string: '{text}'. Not in Settings.tsv.", code_context)

    def is_typo(self, val, target):
        return difflib.SequenceMatcher(None, val.lower(), target.lower()).ratio() > 0.9

    def looks_like_category(self, val, category_set):
        if val in category_set: return True
        for item in category_set:
            if self.is_typo(val, item):
                return True
        return False

    def add_violation(self, file, line, severity, msg, context):
        self.violations.append({
            'file': file,
            'line': line,
            'severity': severity,
            'msg': msg,
            'context': context.strip()
        })

    def generate_report(self):
        print("\n" + "="*60)
        print("       🚀 K&L CRM CODEBASE AUDIT REPORT       ")
        print("="*60 + "\n")

        # 1. Violations
        if self.violations:
            print(f"🚨 FOUND {len(self.violations)} POTENTIAL ISSUES:\n")
            self.violations.sort(key=lambda x: (x['file'], x['line']))
            
            for v in self.violations:
                icon = "🔴" if v['severity'] == "HIGH" else "⚠️"
                print(f"{icon} [{v['severity']}] {v['file']}:{v['line']}")
                print(f"   Issue: {v['msg']}")
                print(f"   Code:  {v['context']}\n")
        else:
            print("✅ No obvious Schema or Settings violations found in string literals.")

        print("-" * 60)
        
        # 2. Function Inventory
        print("\n📂 FUNCTION INVENTORY (By File):")
        for f, funcs in sorted(self.function_map.items()):
            print(f"\n   📄 {f}")
            for func in funcs:
                print(f"      └─ {func}")

        print("\n" + "="*60)
        print("Audit Complete.")

if __name__ == "__main__":
    auditor = CRMAuditor()
    auditor.load_config()
    auditor.scan_codebase()
    auditor.generate_report()