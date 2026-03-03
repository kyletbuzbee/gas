import os
import re
import json
from collections import defaultdict

class KL_CRM_Analyzer:
    def __init__(self, root_dir):
        self.root_dir = root_dir
        self.files = [f for f in os.listdir(root_dir) if f.endswith(('.js', '.gs', '.html', '.csv', '.json'))]
        self.logic_groups = {
            "CORE_ENGINE": ["Config.js", "SharedUtils.js", "DataHelpers.js"],
            "SERVICES": ["ProspectFunctions.js", "OutreachFunctions.js", "AccountFunction.js", "ProspectScoringService.js"],
            "AUTOMATION": ["WorkflowAutomationService.js", "CRMEngine.js"],
            "UI_BACKEND": ["DashboardBackend.gs", "MenuFunctions.js"]
        }
        self.dependency_map = defaultdict(list)

    def analyze_dependencies(self):
        """Maps which files call functions in other files."""
        for filename in self.files:
            if not filename.endswith(('.js', '.gs')): continue
            path = os.path.join(self.root_dir, filename)
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
                # Find calls to global namespaces
                for namespace in ["SharedUtils", "CONFIG", "ProspectFunctions", "OutreachFunctions", "ValidationUtils"]:
                    if namespace in content and namespace not in filename:
                        self.dependency_map[filename].append(namespace)

    def generate_full_report(self):
        """Generates a comprehensive manifest of the CRM ecosystem."""
        manifest_path = "KL_CRM_MANIFEST.txt"
        with open(manifest_path, 'w', encoding='utf-8') as report:
            report.write("=== K&L RECYCLING CRM SYSTEM MANIFEST ===\n")
            report.write(f"Total Files Analyzed: {len(self.files)}\n\n")

            report.write("--- LOGIC GROUPINGS ---\n")
            for group, files in self.logic_groups.items():
                report.write(f"[{group}]: {', '.join(files)}\n")

            report.write("\n--- DEPENDENCY GRAPH (Who calls whom) ---\n")
            for file, deps in self.dependency_map.items():
                report.write(f"{file} relies on: {list(set(deps))}\n")

            report.write("\n--- CONSOLIDATED CODEBASE ---\n")
            for filename in sorted(self.files):
                path = os.path.join(self.root_dir, filename)
                report.write(f"\n\n{'='*50}\n")
                report.write(f"FILE: {filename}\n")
                report.write(f"{'='*50}\n")
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        report.write(f.read())
                except Exception as e:
                    report.write(f"[Error reading file: {e}]")

        print(f"Analysis complete. Report generated: {manifest_path}")

if __name__ == "__main__":
    # Ensure this points to the directory containing your extracted zip files
    analyzer = KL_CRM_Analyzer(os.getcwd())
    analyzer.analyze_dependencies()
    analyzer.generate_full_report()