import os
import pandas as pd

# ---------------------------------
# Step 1: Get Excel file name
# ---------------------------------
file_name = input("Enter Excel file name (example: sample.xlsx): ").strip()

# Current directory
current_directory = os.getcwd()

# Full file path
file_path = os.path.join(current_directory, file_name)

# Check if file exists
if not os.path.exists(file_path):
    print(f"\nFile '{file_name}' not found in current directory.")
    print(f"Current Directory: {current_directory}")
    exit()

# ---------------------------------
# Step 2: Read Excel file
# ---------------------------------
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# ---------------------------------
# Step 3: Print columns and values
# ---------------------------------
print("\n========== EXCEL DATA ==========\n")

for column in df.columns:

    print(f"Column Name: {column}")
    print("Values:")

    for value in df[column]:
        print(f" - {value}")

    print("\n" + "-" * 40)
