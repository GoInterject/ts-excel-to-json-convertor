import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';

// Define types for the structure of the JSON data
interface ExcelSheetData {
  [sheetName: string]: Array<Record<string, any>>;  // Each sheet is an array of row objects
}

export default class ExcelConvertor {

  async chooseConvertor(
    conoversionDirection: string, 
    sourcePath: string, 
    destinationPath: string
  ): Promise<void> {
    const [praparedSourcePath, praparedDestinationPath] = this.preparesSourceAndDestionationPaths(conoversionDirection, sourcePath, destinationPath);
    console.log("preparesSourceAndDestionationPaths received: ", praparedSourcePath, praparedDestinationPath);
    if (conoversionDirection === "convertexcel") {
      this.convertExcelToJSON(praparedSourcePath, praparedDestinationPath)
    } else if (conoversionDirection === "convertjson") {
      this.convertJSONToExcel(praparedSourcePath, praparedDestinationPath)
    }
  }

  preparesSourceAndDestionationPaths(
    conoversionDirection: string, 
    sourcePath: string, 
    destinationPath: string
  ): [string, string] {
    sourcePath = path.resolve(sourcePath); // To get the absolute path
    
    // Check if source file exists
    if (!fs.existsSync(sourcePath)) {
      console.error(`Error: Cannot access file ${sourcePath}. Please check the file path.`);
      process.exit(1);  // Exit the program if file doesn't exist
    }

    const sourceDirPath = path.dirname(sourcePath);
    if (destinationPath && !path.isAbsolute(destinationPath)) {
      destinationPath = path.join(sourceDirPath, destinationPath);
    }
    if (!destinationPath) {
      destinationPath = sourceDirPath;
    }

    // Ensure the output directory exists
    if (!fs.existsSync(destinationPath)) {
      fs.mkdirSync(destinationPath, { recursive: true });
      console.log(`Created output directory: "${destinationPath}"`);
    }

    const stats = fs.statSync(destinationPath);
    if (stats.isDirectory()) {
      // Define output file path
      const sourceFileName = path.basename(sourcePath, path.extname(sourcePath));

      let outputFileExtension: string = ""
      if (conoversionDirection === "convertexcel") {
        outputFileExtension = '.json'
      } else if (conoversionDirection === "convertjson") {
        outputFileExtension = '.xlsx'
      }
      const outputFileName = sourceFileName + outputFileExtension
      destinationPath = path.join(destinationPath, outputFileName);  
    } else if (stats.isFile()) {
      destinationPath = destinationPath;
    } else {
      console.log(`${destinationPath} is neither a file nor a directory.`);
    }

    return [sourcePath, destinationPath];
  }

  async convertExcelToJSON(excelFilePath: string, outputJsonPath: string): Promise<void> {

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
      console.log('\n', `Excel data has been successfully converted and saved to: ${outputJsonPath}`, '\n');
    } catch (err) {
      console.error('Error reading file:', err);
    }
  
  }


  async convertJSONToExcel(inputJsonPath: string, outputExcelPath: string): Promise<void> {

    try {
      // Read the JSON file
      const jsonData = fs.readFileSync(inputJsonPath, 'utf-8');
      
      // Parse the JSON data
      const parsedJSONData: Record<string, any> = JSON.parse(jsonData);
    
      // Create a new workbook
      const workbook: XLSX.WorkBook = XLSX.utils.book_new();
    
      // Iterate through each sheet in the JSON data
      for (const sheetName in parsedJSONData) {
        if (parsedJSONData.hasOwnProperty(sheetName)) {
          const sheetData = parsedJSONData[sheetName];
    
          // Convert each sheet's data (array of objects) into a sheet in the Excel workbook
          const sheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(sheetData);
          
          // Add the sheet to the workbook
          XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
        }
      }

      const options: XLSX.WritingOptions = {
        bookType: 'xlsx', // Specify the Excel format (xlsx, xlsm, etc.)
        type: 'buffer',   // Specifies the output format
      };
      const excelData = XLSX.write(workbook, options)
      fs.writeFileSync(outputExcelPath, excelData)

      console.log('\n', `JSON data has been successfully converted and saved to: ${outputExcelPath}`, '\n');
    } catch (err) {
      console.error('Error reading file:', err);
    }
  }

}


const conversionDirection: string = process.argv[2];
const sourcePath: string = process.argv[3];
const destinationPath: string = process.argv[4];

const excel_convertor = new ExcelConvertor()
Promise.resolve(excel_convertor.chooseConvertor(conversionDirection, sourcePath, destinationPath));
