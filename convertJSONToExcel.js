const fs = require("fs");
const XLSX = require("xlsx");

const workbook = XLSX.utils.book_new();
const excelFileName = "output.xlsx";


// Function to create a hyperlink formula for Excel
function createHyperlinkFormula(sheetName) {
    return `=HYPERLINK("#'${sheetName}'!A1", "SHEET::${sheetName}")`;
}

// Processing the Arrays
function processArray(jsonData, SheetName) {
    let Data = [];
    let cellFormula = "";
    Data[0] = [];
    try{
    for (let i = 0; i < jsonData.length; i++) {
        if (jsonData[i] === null) {
            cellFormula += " ";
        } else if (Array.isArray(jsonData[i])) {
            const nestedSheetName = `${SheetName}.${i}`;
            cellFormula += createHyperlinkFormula(nestedSheetName);
            processArray(jsonData[i], nestedSheetName);
        } else if (typeof jsonData[i] === "object") {
            const nestedSheetName = `${SheetName}.${i}`;
            cellFormula += createHyperlinkFormula(nestedSheetName);
            ProcessObjects(jsonData[i], nestedSheetName);
        } else {
            cellFormula += jsonData[i];
        }

        cellFormula += ", ";
    }
    Data[0].push(cellFormula);
    XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(Data),
        SheetName
    );

    XLSX.writeFile(workbook, excelFileName);
    } catch (error) {
        console.error(`Error processing array in sheet '${sheetName}':`, error.message);
    }
}

// Function to process objects in JSON
function ProcessObjects(jsonData, SheetName) {
    let Data = [];
    Data[0] = Object.keys(jsonData);

    try{

    if (typeof jsonData === "object") {
        Data[1] = [];
        for (const key in jsonData) {
            if (
                typeof jsonData[key] === "object" &&
                jsonData[key] != null &&
                !Array.isArray(jsonData[key])
            ) {
                const nestedSheetName = `${SheetName}.${key}`;
                ProcessObjects(jsonData[key], nestedSheetName);
                Data[1].push(createHyperlinkFormula(nestedSheetName));
            } else if (Array.isArray(jsonData[key])) {
                let cellFormula = "";
                for (let i = 0; i < jsonData[key].length; i++) {
                    if (jsonData[key][i] === null) {
                        cellFormula += " ";
                    } else if (Array.isArray(jsonData[key][i])) {
                        const nestedSheetName = `${SheetName}.${key}.${i}`;
                        cellFormula += createHyperlinkFormula(nestedSheetName);
                        processArray(jsonData[key][i], nestedSheetName);
                    } else if (typeof jsonData[key][i] === "object") {
                        const nestedSheetName = `${SheetName}.${key}.${i}`;
                        cellFormula += createHyperlinkFormula(nestedSheetName);
                        ProcessObjects(jsonData[key][i], nestedSheetName);
                    } else {
                        cellFormula += jsonData[key][i];
                    }

                    cellFormula += ", ";
                }
                Data[1].push(cellFormula);
            } else {
                Data[1].push(jsonData[key]);
            }
        }
    } else if (Array.isArray(jsonData)) {
        processArray(jsonData, SheetName);
    } else {
        console.log("JsonData is empty");
        return;
    }

    XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet(Data),
        SheetName
    );

    XLSX.writeFile(workbook, excelFileName);
    }catch(error){
        console.error(`Error processing objects in sheet '${sheetName}':`, error.message);
    }
}

// Function to read and process JSON from a file
function readAndProcessJSON(JsonPath) {
    try {
        const JsonData = JSON.parse(fs.readFileSync(JsonPath, "utf-8"));
        // Function to Process JsonObjects
        ProcessObjects(JsonData, "Sheet1");
        console.log("Conversion successful. Excel file created:", excelFileName);
    } catch (error) {
        console.error("Error parocessing JSON:", error.message);
    }
}

// Path of the File
const JsonPath = "file.json";

// Data read synchronously and then parsed for JSON-formatted String to Javascript Object
readAndProcessJSON(JsonPath);
