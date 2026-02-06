import pandas as pd

# Load the CSV file
file_path = "/Users/brandon/Downloads/lpp-export-fbcbc03d-d80e-4315-a46e-ee60e3117c49.csv"  # <-- replace with your CSV file path
df = pd.read_csv(file_path)

# Count occurrences of each distinct buyer
buyer_counts = df['Buyer'].value_counts()

# Filter out counts less than 3
buyer_counts_filtered = buyer_counts[buyer_counts >= 2]

# Reset index to turn Buyer names into a column
buyer_counts_filtered = buyer_counts_filtered.reset_index()
buyer_counts_filtered.columns = ['Buyer', 'Count']

# Add row numbers starting from 1
buyer_counts_filtered.insert(0, 'RowNumber', range(1, len(buyer_counts_filtered) + 1))

# Add OwnerType (most common OwnerType for that buyer)
owner_type_list = []
for buyer in buyer_counts_filtered['Buyer']:
    most_common_owner_type = df[df['Buyer'] == buyer]['OwnerType'].mode()
    if not most_common_owner_type.empty:
        owner_type_list.append(most_common_owner_type[0])
    else:
        owner_type_list.append(None)

buyer_counts_filtered['OwnerType'] = owner_type_list

# Display the results
print(buyer_counts_filtered)

# Save to a new CSV file
buyer_counts_filtered.to_csv("buyer_counts_filtered_with_owner_type.csv", index=False)
