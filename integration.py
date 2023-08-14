import openai
import pandas as pd

# Set your OpenAI API key here
openai.api_key = "YOUR_OPENAI_API_KEY"

# Read the Excel file
excel_file = "path_to_your_excel_file.xlsx"
sheet_name = "Sheet1"  # Change this to your sheet name
df = pd.read_excel(excel_file, sheet_name=sheet_name)

descriptions = df["Column A"]  # Assuming descriptions are in Column A

# Initialize an empty table for storing the extracted information
result_table = []

# Process descriptions using GPT-3 and extract relevant information
for description in descriptions:
    prompt = f"Extract the type of equipment, model, and brand from the following description:\n{description}"
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        max_tokens=100,
        stop=None,
        temperature=0,
    )
    
    extracted_info = response.choices[0].text.strip()
    info_lines = extracted_info.split("\n")
    info_dict = {}
    for line in info_lines:
        key, value = map(str.strip, line.split(":"))
        info_dict[key] = value
    
    result_table.append(info_dict)

# Create a new DataFrame from the extracted information
result_df = pd.DataFrame(result_table, columns=["Type of Equipment", "Model", "Brand"])

# Save the new DataFrame to a new Excel file
result_excel_file = "extracted_info.xlsx"
result_df.to_excel(result_excel_file, index=False)
print(f"Extracted information saved to {result_excel_file}")
