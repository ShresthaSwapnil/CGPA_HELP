import pandas as pd

# Load your Excel file
file_path = "Marks.xlsx"  # Ensure this file is in the same directory
xls = pd.ExcelFile(file_path)

# Load the first sheet
df = xls.parse(xls.sheet_names[0])

# Rename columns
df.columns = ['Subject Code', 'Subject Name', 'Grade']

# Add Serial Number (S.N.)
df.insert(0, 'S.N.', range(1, len(df) + 1))

# Print the full subject list
print("=== Subject List ===")
print(df.to_string(index=False))

# Mapping of grades to grade points
grade_point_map = {
    "HD1": 4.00,
    "HD2": 3.75,
    "DI1": 3.50,
    "DI2": 3.25,
    "CR1": 3.00,
    "CR2": 2.75,
    "PS1": 2.50,
    "PS2": 2.00,
    "FL": 0.00
}

# Extract grades and map to grade points
grades = df['Grade']
grade_points = grades.map(grade_point_map)

# Calculate CGPA
cgpa = grade_points.mean()

# Show formula used
print("\nFormula used:")
print("CGPA = (Sum of all Grade Points) / (Number of Subjects)")
print(f"CGPA = ({' + '.join(f'{gp:.2f}' for gp in grade_points if pd.notnull(gp))}) / {len(grade_points)}")

# Optional: Degree Classification
if cgpa >= 3.75:
    classification = "First Class"
elif cgpa >= 3.25:
    classification = "Second Upper Class"
elif cgpa >= 2.75:
    classification = "Second Lower Class"
elif cgpa >= 2.50:
    classification = "Third Class"
elif cgpa >= 2.00:
    classification = "General Degree"
else:
    classification = "Below Degree Level"


print(f"\nYour CGPA is: {cgpa:.2f}")

print(f"\nDegree Classification: {classification}")
