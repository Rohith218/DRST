import pandas as pd

# Sample data
data = {
    "A": ["Marvel", "Ironman", "Spiderman", "Hulk"],
    "B": ["DC", "Batman", "Superman", "Flash"],
    "C": ["Raghu", "Raghav", "Ramesh", "Rahul"],
    "D": ["Hyd", "Chennai", "Delhi", "Mumbai"],
    "K": ["Hyd", "Chennai", "Delhi", "Mumbai"]
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
file_name = "sample1001.xlsx"
df.to_excel(file_name, index=False)

print(f"Excel file '{file_name}' created successfully.")

