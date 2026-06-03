import os

from SearchEngine import SearchEngine


def format_path(file_path, root_dir):

    relative_path = os.path.relpath(
        file_path,
        root_dir
    )

    return " → ".join(
        relative_path.split(os.sep)
    )


# ================= FOLDER INPUT =================

folder = input(
    "Enter folder path: "
).strip()

if not os.path.exists(folder):

    print("\nFolder does not exist.\n")

    exit()

# ================= LOAD ENGINE =================

print("\nLoading Excel files...\n")

engine = SearchEngine(folder)

files = engine.files

if not files:

    print("No Excel files found.")

    exit()

# ================= DISPLAY FILES =================

print("Excel files detected:\n")

for f in files:

    print(
        "-",
        format_path(f, folder)
    )

# ================= LOAD COLUMNS =================

print("\nLoading columns...\n")

columns = engine.get_all_columns()

if not columns:

    print("No columns detected.")

    exit()

print("Detected Columns:\n")

for i, col in enumerate(columns, start=1):

    print(f"{i}. {col}")

# ================= CONDITIONS =================

try:

    num_checks = int(
        input("\nHow many conditions? ")
    )

except ValueError:

    print("Invalid number.")

    exit()

checks = {}

for i in range(num_checks):

    col = input(
        f"\nColumn {i+1}: "
    ).strip()

    val = input(
        f"Value for '{col}': "
    ).strip()

    checks[col] = val

# ================= SEARCH =================

print("\nScanning Excel files...\n")


def progress(current, total):

    percent = int(
        (current / total) * 100
    )

    print(
        f"\rProgress: {percent}% ({current}/{total})",
        end=""
    )


summary_results, final_df = engine.run_search(
    checks,
    progress
)

# ================= RESULTS =================

print("\n\n========== RESULTS ==========\n")

if not summary_results:

    print("No matches found.")

    exit()

total_matches = 0

for file, result in summary_results:

    path = format_path(
        file,
        folder
    )

    print(
        f"{path} → {result} match(es)"
    )

    if isinstance(result, int):
        total_matches += result

print(
    f"\nTotal Matched Rows: {total_matches}"
)

# ================= EXPORT =================

if final_df.empty:

    print(
        "\nNo rows available for export."
    )

    exit()

choice = input(
    "\nExport matched rows to Excel? (y/n): "
).strip().lower()

if choice != "y":

    print("\nExport skipped.")

    exit()

output_path = input(
    "\nEnter output Excel filename/path: "
).strip()

if not output_path.endswith(".xlsx"):

    output_path += ".xlsx"

try:

    engine.export_results(
        final_df,
        output_path
    )

    print(
        f"\nExcel exported successfully:\n{output_path}"
    )

except Exception as e:

    print(
        f"\nExport failed:\n{e}"
    )