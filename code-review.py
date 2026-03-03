#!/usr/bin/env python3
"""
AI-Optimized Grouped File Converter
Consolidates source code into type-specific .txt files with high-visibility AI markers.
"""

import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

class AIOptimizedGroupedConverter:
    def __init__(self, source_folder, output_dir):
        self.source_folder = Path(source_folder)
        self.output_dir = Path(output_dir)
        self.extensions = ['.html', '.json', '.gs', '.css', '.yml', '.yaml', '.ts', '.tsx', '.jsx', '.java', '.c', '.cpp', '.rb', '.go', '.rs', '.ts', '.env.local', '.env', '.css', '.scss', '.csv', '.xml', '.gs', '.js']
        self.exclude_dirs = ['node_modules', '.git', 'dist', 'build', 'coverage', '.qodo', 'ai_logic_review', '__pycache__']
        self.stats = {"converted": 0, "groups": {}}

    def format_ai_separator(self, file_path):
        """
        Creates a high-contrast separator for AI processing.
        Includes full path and metadata for context grounding.
        """
        rel_path = file_path.relative_to(self.source_folder)
        return (
            f"\n{'=' * 80}\n"
            f"FILE_BEGIN: {rel_path}\n"
            f"METADATA: Size={file_path.stat().st_size} bytes | Last_Modified={datetime.fromtimestamp(file_path.stat().st_mtime)}\n"
            f"{'=' * 80}\n"
        )

    def format_eof_marker(self, file_path):
        """Adds a clear boundary marker for the end of the file content."""
        rel_path = file_path.relative_to(self.source_folder)
        return f"\n\n[FILE_END: {rel_path}]\n{'#' * 80}\n"

    def process_files(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)
        all_files = []
        for root, dirs, files in os.walk(self.source_folder):
            dirs[:] = [d for d in dirs if d not in self.exclude_dirs]
            for f in files:
                all_files.append(Path(root) / f)
        
        all_files.sort()

        for f_path in all_files:
            if f_path.suffix.lower() not in self.extensions:
                continue

            # Group by extension: e.g., js_logic.txt, html_logic.txt
            ext_key = f_path.suffix.lower().replace('.', '')
            output_file = self.output_dir / f"{ext_key}_logic_group.txt"
            
            try:
                with open(f_path, 'r', encoding='utf-8', errors='ignore') as src:
                    content = src.read()
                
                if not content.strip():
                    continue

                with open(output_file, 'a', encoding='utf-8') as out:
                    # Write high-visibility BEGIN marker
                    out.write(self.format_ai_separator(f_path))
                    out.write(content)
                    # Write clear EOF marker
                    out.write(self.format_eof_marker(f_path))

                self.stats['converted'] += 1
                self.stats['groups'][ext_key] = self.stats['groups'].get(ext_key, 0) + 1
                print(f"[OK] {f_path.name} appended to {output_file.name}")
            except Exception as e:
                print(f"[ERR] {f_path.name}: {e}")

    def run(self):
        print(f"Exporting AI-ready groups from: {self.source_folder}")
        self.process_files()
        print("\n--- Summary ---")
        for group, count in self.stats['groups'].items():
            print(f"Type {group.upper()}: {count} files -> {group}_logic_group.txt")

if __name__ == "__main__":
    # Scans current directory and outputs to 'ai_logic_review' folder
    converter = AIOptimizedGroupedConverter('.', 'ai_logic_review')
    converter.run()