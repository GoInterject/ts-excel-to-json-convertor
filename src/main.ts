import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';

// Define types for the structure of the JSON data
interface ExcelSheetData {
  [sheetName: string]: Array<Record<string, any>>;  // Each sheet is an array of row objects
}

export default class ExcelConvertor {

  async convertExcelToJSON(excelFilePath: string, outputDir: string): Promise<void> {
    // Check if Excel file exists
    if (!fs.existsSync(excelFilePath)) {
      console.error(`Error: Cannot access file ${excelFilePath}. Please check the file path.`);
      process.exit(1);  // Exit the program if file doesn't exist
    }
  
    if (outputDir && !path.isAbsolute(outputDir)) {
      const excelDirPath = path.dirname(excelFilePath);
      outputDir = path.join(excelDirPath, outputDir);
    }
    if (!outputDir) {
      outputDir = path.dirname(excelFilePath);
    }

    // Ensure the output directory exists
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log(`Created output directory: "${outputDir}"`);
    }
    
    let outputJsonPath: string = "";
    const stats = fs.statSync(outputDir);
    if (stats.isDirectory()) {
      // Define output JSON file path
      const excelFileName = path.basename(excelFilePath, path.extname(excelFilePath));
      const jsonFileName = excelFileName + '.json'
      outputJsonPath = path.join(outputDir, jsonFileName);  
    } else if (stats.isFile()) {
      outputJsonPath = outputDir;
    } else {
      console.log(`${outputDir} is neither a file nor a directory.`);
    }

    try {
      const data = fs.readFileSync(excelFilePath);
    
      const workbook: XLSX.WorkBook = XLSX.read(data);
    
      // Convert workbook to JSON (each sheet as an array of rows)
      const sheetData: ExcelSheetData = workbook.SheetNames.reduce((acc: ExcelSheetData, sheetName: string) => {
        const sheet: XLSX.WorkSheet = workbook.Sheets[sheetName];
        const jsonData: Array<Record<string, any>> = XLSX.utils.sheet_to_json(sheet);
        acc[sheetName] = jsonData;
        return acc;
      }, {});
    
      // Write JSON data to the output file
      fs.writeFileSync(outputJsonPath, JSON.stringify(sheetData, null, 2), 'utf-8');
      console.log(`Excel data has been successfully converted and saved to: ${outputJsonPath}`);
    } catch (err) {
      console.error('Error reading file:', err);
    }
  
  }
}


const excelFilePath: string = process.argv[2];
let outputDir: string = process.argv[3];

const excel_convertor = new ExcelConvertor()
Promise.resolve(excel_convertor.convertExcelToJSON(excelFilePath, outputDir));
