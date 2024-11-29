const Excel = require("exceljs");
const ifsc = require("ifsc");
const readline = require("readline");
const axios = require("axios");

const workbook = new Excel.Workbook();
const bank_list = {}; // Stores valid bank details by IFSC code
const valid_list = {}; // Stores validation results for IFSC codes
const lookupHistory = []; // Stores IFSC lookup history

// Function to process the Excel file
async function processExcelFile() {
  try {
    await workbook.xlsx.readFile("sample.xlsx"); // Read the Excel file
    const worksheet = workbook.getWorksheet(1); // Get the first worksheet
    const total = worksheet.rowCount; // Get total number of rows
    let count = 0,
      valid = 0,
      bar = "█";

    for (let i = 1; i <= total; i++) {
      const row = worksheet.getRow(i);
      const code = row.getCell(1).value;

      if (!(code in valid_list)) valid_list[code] = ifsc.validate(code); // Validate IFSC code

      if (valid_list[code]) {
        if (!(code in bank_list)) bank_list[code] = await ifsc.fetchDetails(code); // Fetch bank details if valid

        const details = bank_list[code];
        row.getCell(2).value = details.BANK;
        row.getCell(3).value = details.BRANCH;
        valid++;

        // Record the lookup history for valid codes
        lookupHistory.push({
          ifsc: code,
          bank: details.BANK,
          branch: details.BRANCH
        });
      } else {
        row.getCell(2).value = "Invalid IFSC";
        row.getCell(2).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0000" },
        };
      }
      count++;
      if (count % 20 === 0) bar += "█"; // Display progress bar
      console.clear();
      console.log(`\n${bar} \nProcessed ${count} out of ${total}`);
    }

    console.log(`\nValid: ${valid}\nInvalid: ${total - valid}`);
    console.log("\nDONE...");

    // Save updated Excel file with bank details
    await workbook.xlsx.writeFile("output.xlsx");

    // After processing, prompt the user for region input
    promptRegionInput();

  } catch (error) {
    console.error("Error processing the Excel file:", error.message);
  }
}

// Function to prompt the user for region input
function promptRegionInput() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  rl.question("\nEnter the region name (state) to fetch bank details: ", async (regionName) => {
    if (!regionName || regionName.trim() === "") {
      console.log("\nNo region provided. Please try again.");
      rl.close();
      promptRegionInput(); // Retry asking for the region
    } else {
      await getBanksByLocation(regionName.trim().toLowerCase()); // Fetch data for the entered region
      rl.close();

      // After fetching region banks, ask if the user wants to see the lookup history
      promptLookupHistory();
    }
  });
}

// Function to fetch bank details based on location using OpenStreetMap
async function getBanksByLocation(location) {
  const url = `https://nominatim.openstreetmap.org/search?q=bank+in+${encodeURIComponent(location)}&format=json&addressdetails=1`;

  try {
    const response = await axios.get(url);
    const results = response.data;

    if (results && results.length > 0) {
      console.log(`Banks in ${location}:`);
      for (let place of results) {
        const bankName = place.display_name; // Display name of the bank
        const address = place.address
          ? `${place.address.road || ""}, ${place.address.city || ""}, ${place.address.state || ""}, ${place.address.country || ""}`
          : "Address not available";

        console.log(`Name: ${bankName}`);
        console.log(`Address: ${address}`);

        console.log('----------------------');
      }
    } else {
      console.log(`No banks found for ${location}.`);
    }
  } catch (error) {
    console.error("Error fetching bank details:", error.message);
  }
}

// Function to prompt user for IFSC code lookup history
function promptLookupHistory() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  rl.question("\nWould you like to view the IFSC code lookup history? (yes/no): ", (answer) => {
    if (answer.toLowerCase() === 'yes') {
      if (lookupHistory.length > 0) {
        console.log("\nIFSC Code Lookup History:");
        lookupHistory.forEach((entry, index) => {
          console.log(`${index + 1}. IFSC: ${entry.ifsc}, Bank: ${entry.bank}, Branch: ${entry.branch}`);
        });
      } else {
        console.log("\nNo lookup history found.");
      }
    } else {
      console.log("\nExiting...");
    }
    rl.close();
  });
}

// Start the process
processExcelFile();
