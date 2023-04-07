import pandas as pd

# Read the Excel file
df = pd.read_excel("family_trees.xlsx")

# Find the separators and split the data into family trees
separators = df[df.iloc[:,0] == "SEPERATOR"].index.tolist()
separators.append(len(df))
family_trees = []
start = 0
for sep in separators:
    family_trees.append(df.iloc[start:sep, :])
    start = sep+1

# Process each family tree
for i, tree in enumerate(family_trees):
    print(f"\nFamily Tree {i+1}:")
    # Create a dictionary to hold the family tree
    family_dict = {}
    for _, row in tree.iterrows():
        # Skip the separator and empty rows
        if row.iloc[0] in ["SEPERATOR", None]:
            continue
        # Find the parent tag and the child tag
        parent_tag = row.iloc[0]
        for j in range(1, len(row)):
            child_tag = row.iloc[j]
            if pd.isna(child_tag):
                break
            # Add the child tag to the parent's list of children
            if parent_tag in family_dict:
                family_dict[parent_tag].append(child_tag)
            else:
                family_dict[parent_tag] = [child_tag]
    # Print the family tree
    def print_tree(tag, indent):
        children = family_dict.get(tag, [])
        if not children:
            return
        for i, child in enumerate(children):
            if i == len(children) - 1:
                print("  "*indent + "└─ " + child)
                print_tree(child, indent+2)
            else:
                print("  "*indent + "├─ " + child)
                print_tree(child, indent+2)
    print_tree(tree.iloc[0,0], 0)
