import os
import re
from typing import Optional


class KL_CRM_Mapper:
    """Analyzes K&L CRM JavaScript/Google Apps Script files for dependencies and architectural structure."""

    # Configuration constants - follows Config-First principle
    SUPPORTED_EXTENSIONS = ('.js', '.gs', '.html')
    DEFAULT_OUTPUT_FILE = 'CRM_FULL_REPORT.txt'

    # Logic groups defining the 4-tier architecture (verified against actual files)
    LOGIC_GROUPS: dict[str, list[str]] = {
        "DATA_ENGINE": ["SharedUtils.js", "DataHelpers.js", "Config.js", "Normalization.js", "DataValidation.js", "BusinessValidation.js", "CSVImport.js", "SchemaNormalizer.js"],
        "BUSINESS_LOGIC": ["ProspectFunctions.js", "OutreachFunctions.js", "AccountFunction.js", "ProspectScoringService.js", "SalesFunctions.js", "ActiveContainerFunctions.js"],
        "VALIDATION": ["ValidationUtils.js", "BusinessValidation.js", "DataValidation.js", "ErrorHandling.js", "FuzzyMatchingUtils.js", "ComprehensiveValidationSystem.js", "SettingsValidation.js"],
        "UI_BACKEND": ["DashboardBackend.js", "MenuFunctions.js", "ReportFunctions.js", "CRMBrain.js"],
        "WORKFLOW_AUTOMATION": ["Sync.js", "PipelineService.js", "WorkflowAutomationService.js", "BatchProcessor.js"],
    }

    # Compiled regex for dependency detection - compiled once for performance
    _DEPENDENCY_PATTERN = re.compile(r'(SharedUtils|CONFIG|ProspectFunctions|OutreachFunctions|ValidationUtils)\.')

    def __init__(
        self,
        target_dir: str = ".",
        output_file: Optional[str] = None,
        extensions: Optional[tuple[str, ...]] = None,
        logic_groups: Optional[dict[str, list[str]]] = None
    ):
        """
        Initialize the CRM logic mapper.

        Args:
            target_dir: Directory to scan for files (default: current directory)
            output_file: Output report filename (default: CRM_FULL_REPORT.txt)
            extensions: File extensions to scan (default: .js, .gs, .html)
            logic_groups: Custom logic groups for analysis (default: class constants)
        """
        self.target_dir = target_dir
        self.output_file = output_file or self.DEFAULT_OUTPUT_FILE
        self.extensions = extensions or self.SUPPORTED_EXTENSIONS
        self.logic_groups = logic_groups or self.LOGIC_GROUPS

        # Discover files and cache for content reuse
        self.files = self._discover_files()
        self._content_cache: dict[str, str] = {}

    def _discover_files(self) -> list[str]:
        """Discover all relevant files in the target directory."""
        try:
            return [
                f for f in os.listdir(self.target_dir)
                if f.endswith(self.extensions)
            ]
        except OSError as e:
            print(f"[ERROR] Cannot access directory '{self.target_dir}': {e}")
            return []

    def _read_file(self, filename: str) -> Optional[str]:
        """
        Read file content with caching to avoid redundant I/O.

        Args:
            filename: Name of the file to read

        Returns:
            File content as string, or None if read fails
        """
        # Return cached content if available
        if filename in self._content_cache:
            return self._content_cache[filename]

        filepath = os.path.join(self.target_dir, filename)
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
                self._content_cache[filename] = content
                return content
        except FileNotFoundError:
            self._report_warning(f"File not found: {filename}")
        except PermissionError:
            self._report_warning(f"Permission denied: {filename}")
        except UnicodeDecodeError as e:
            self._report_warning(f"Encoding error in {filename}: {e}")
        except OSError as e:
            self._report_warning(f"Error reading {filename}: {e}")

        return None

    def _report_warning(self, message: str) -> None:
        """Log a warning message."""
        print(f"[WARNING] {message}")

    def _extract_dependencies(self, content: str) -> list[str]:
        """
        Extract dependencies from file content using compiled regex.

        Args:
            content: File content to analyze

        Returns:
            List of unique dependencies found
        """
        matches = self._DEPENDENCY_PATTERN.findall(content)
        return list(set(matches)) if matches else []

    def analyze(self) -> dict[str, any]:
        """
        Perform the complete CRM logic analysis.

        Returns:
            Analysis result dictionary with success status and data
        """
        if not self.files:
            return {
                "success": False,
                "data": None,
                "error": f"No files found with extensions {self.extensions} in {self.target_dir}"
            }

        report = []

        # Header
        report.append("=== K&L RECYCLING CRM LOGIC ANALYSIS ===\n")
        report.append(f"Scanned directory: {self.target_dir}")
        report.append(f"Files found: {len(self.files)}\n")

        # 1. Map Connections - Cross-File Dependencies
        report.append("--- CROSS-FILE DEPENDENCIES ---")
        dependencies_found = False
        for filename in self.files:
            content = self._read_file(filename)
            if content is None:
                continue

            deps = self._extract_dependencies(content)
            if deps:
                report.append(f"{filename} calls: {deps}")
                dependencies_found = True

        if not dependencies_found:
            report.append("[INFO] No cross-file dependencies detected.")

        # 2. Identify the _rowIndex Bug
        report.append("\n--- CRITICAL BUG DETECTION: _rowIndex Injection ---")
        sharedutils_content = self._read_file('SharedUtils.js')
        if sharedutils_content is not None:
            if '_rowIndex' in sharedutils_content:
                report.append("[PASS] SharedUtils.js contains _rowIndex injection.")
            else:
                report.append("[FAIL] SharedUtils.js is MISSING _rowIndex injection. This causes the 'undefined' errors.")
        else:
            report.append("[WARNING] SharedUtils.js not found in current directory.")

        # 3. Generate Consolidated Codebase
        report.append("\n--- CONSOLIDATED LOGIC GROUPS ---")
        for group_name, group_files in self.logic_groups.items():
            report.append(f"\n>> GROUP: {group_name}")
            found_count = 0
            for gf in group_files:
                if gf in self.files:
                    report.append(f"   - {gf}")
                    found_count += 1
            report.append(f"   [{found_count}/{len(group_files)} files present]")

        # Write report to file
        try:
            with open(self.output_file, 'w', encoding='utf-8') as out:
                out.write("\n".join(report))
            print(f"Report generated: {self.output_file}")
        except OSError as e:
            return {
                "success": False,
                "data": None,
                "error": f"Failed to write report: {e}"
            }

        return {
            "success": True,
            "data": {
                "files_analyzed": len(self.files),
                "output_file": self.output_file,
                "logic_groups": list(self.logic_groups.keys())
            },
            "error": None
        }


def main() -> None:
    """Main entry point for the CRM logic mapper."""
    import argparse

    parser = argparse.ArgumentParser(description="K&L CRM Logic Analysis Tool")
    parser.add_argument(
        "-d", "--directory",
        default=".",
        help="Directory to scan (default: current directory)"
    )
    parser.add_argument(
        "-o", "--output",
        default="CRM_FULL_REPORT.txt",
        help="Output report filename"
    )

    args = parser.parse_args()

    mapper = KL_CRM_Mapper(
        target_dir=args.directory,
        output_file=args.output
    )
    result = mapper.analyze()

    if not result["success"]:
        print(f"Error: {result['error']}")
        exit(1)


if __name__ == "__main__":
    main()