import csv
import glob
import sys
from pathlib import Path

OUTPUT_FILE = "combined_headers.txt"


def main():
    csv_files = sorted(glob.glob("*.csv"))

    if not csv_files:
        print("No CSV files found in the current directory.", file=sys.stderr)
        sys.exit(1)

    processed_count = 0
    error_count = 0

    try:
        with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as outfile:
            writer = csv.writer(outfile)

            for csv_file in csv_files:
                file_path = Path(csv_file)

                if not file_path.is_file():
                    print(f"Warning: Skipping '{csv_file}' (not a file)", file=sys.stderr)
                    error_count += 1
                    continue

                try:
                    with open(csv_file, "r", newline="", encoding="utf-8") as infile:
                        reader = csv.reader(infile)

                        try:
                            headers = next(reader)
                        except StopIteration:
                            print(f"Warning: '{csv_file}' is empty (no headers)", file=sys.stderr)
                            error_count += 1
                            continue

                        outfile.write(f"=== {csv_file} ===\n")
                        writer.writerow(headers)
                        processed_count += 1

                except PermissionError:
                    print(f"Error: Permission denied reading '{csv_file}'", file=sys.stderr)
                    error_count += 1
                except UnicodeDecodeError as e:
                    print(f"Error: Encoding issue in '{csv_file}': {e}", file=sys.stderr)
                    error_count += 1
                except csv.Error as e:
                    print(f"Error: CSV parsing error in '{csv_file}': {e}", file=sys.stderr)
                    error_count += 1

    except PermissionError:
        print(f"Error: Permission denied writing to '{OUTPUT_FILE}'", file=sys.stderr)
        sys.exit(1)
    except OSError as e:
        print(f"Error: Cannot write to '{OUTPUT_FILE}': {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Combined headers from {processed_count} CSV files into {OUTPUT_FILE}")
    if error_count > 0:
        print(f"Encountered {error_count} error(s). Check stderr for details.", file=sys.stderr)


if __name__ == "__main__":
    main()
