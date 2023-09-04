// Import the necessary libraries
const openai = require("openai");
const pandas = require("pandas");
const xlwings = require("xlwings");
const os = require("os");

// Set up your GPT-3 API key using an environment variable
openai.api_key = os.getenv("API_KEY");

// Define a function that generates GPT-3 responses for a list of input strings
function generateGpt3Responses(inputList) {
  const outputList = [];
  for (const inputString of inputList) {
    try {
      const response = await openai.Completion.create({
        engine: "text-davinci-002",
        prompt: inputString,
        max_tokens: 100,
        n: 1,
        stop: ["\\n"]
      });
      outputList.push(response.choices[0].text);
    } catch (e) {
      console.log(f"Error generating GPT-3 response: {e}");
      outputList.push("");
    }
  }
  return outputList;
}

// Open the Excel file
const wb = new xlwings.Book("your_file.xlsx");

// Select the worksheet
const ws = wb.sheets["Sheet1"];

// Convert the worksheet to a Pandas DataFrame
const tableString = ws.range("A1").expand("table").value;

const tableList = [];
for (const row of tableString.split("\n")) {
  tableList.push(row.split(","));
}

const df = new pandas.DataFrame(tableList);

// Create the Description column
df["Description"] = "";

// Apply the GPT-3 integration function to the DataFrame
df["GPT-3 Integration"] = generateGtp3Responses(df["Description"]);

// Create a new worksheet with the GPT-3 integration results
const newWs = wb.sheets.add("GPT-3 Integration");

// Copy the GPT-3 integration results to the new worksheet
newWs.range("A1").value = df["GPT-3 Integration"];

// Save the Excel file
wb.save();