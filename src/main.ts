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
    conversionMode: string, 
    sourcePath: string, 
    destinationPath: string
  ): Promise<void> {

    const [praparedSourcePath, praparedDestinationPath] = this.preparesSourceAndDestionationPaths(sourcePath, destinationPath);

    let selectedAction: (conversionMode: string, sourcePath: string, destinationPath: string) => void = () => {};

    if (conoversionDirection === "convertexcel") {
      selectedAction = this.convertExcelToJSON;
    } else if (conoversionDirection === "convertjson") {
      selectedAction = this.convertJSONToExcel;
    }
    
    const statsSourcePath = fs.statSync(praparedSourcePath);
    const statsDestinationPath = fs.statSync(praparedDestinationPath);
    if (statsSourcePath.isDirectory() && statsDestinationPath.isDirectory()) {

      const files = fs.readdirSync(praparedSourcePath);
      files.forEach((file) => {
        const sourceFilePath = path.join(praparedSourcePath, file);
        const stats = fs.statSync(sourceFilePath);
        if (stats.isFile()) {
          const fileExtension = path.extname(file);  // Get the file extension
          if (fileExtension === this.getFileExtFromDirectionConversion(conoversionDirection, 1)) {
            const destinationFilePath: string = this.getDestinationFilePath(
              conoversionDirection, 
              sourceFilePath, 
              praparedDestinationPath
            );
            selectedAction(conversionMode, sourceFilePath, destinationFilePath);
          }
        }
      });

    } else if (statsSourcePath.isFile()) {

      const destinationFilePath: string = this.getDestinationFilePath(
        conoversionDirection, 
        praparedSourcePath, 
        praparedDestinationPath
      );
      selectedAction(conversionMode, praparedSourcePath, destinationFilePath);

    }

  }

  // Created needed folder and returns full paths.
  preparesSourceAndDestionationPaths(
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

    return [sourcePath, destinationPath];
  }

  getFileExtFromDirectionConversion(conoversionDirection: string, inverse: number = 0) {
    const conoversionDirections: Array<string> = ["convertexcel", "convertjson"]
    const fileExtIndex: number = conoversionDirections.indexOf(conoversionDirection) ^ (inverse ^ 1)
    const outputFileExtensions: Array<string> = ['.xlsx', '.json']
    return outputFileExtensions[fileExtIndex]
  }

  // Adds file name output file to the dir path.
  getDestinationFilePath(
    conoversionDirection: string, 
    sourcePath: string, 
    destinationPath: string
  ): string {
    
    const stats = fs.statSync(destinationPath);
    if (stats.isDirectory()) {
      // Define output file path
      const sourceFileName = path.basename(sourcePath, path.extname(sourcePath));

      const outputFileExtension: string = this.getFileExtFromDirectionConversion(conoversionDirection);

      const outputFileName = sourceFileName + outputFileExtension
      destinationPath = path.join(destinationPath, outputFileName);
    } else if (stats.isFile()) {
      destinationPath = destinationPath;
    } else {
      console.log(`${destinationPath} is neither a file nor a directory.`);
    }

    return destinationPath;
  }

  async convertExcelToJSON(conversionMode: string, excelFilePath: string, outputJsonPath: string): Promise<void> {

    try {
      const data = fs.readFileSync(excelFilePath);
    
      const workbook: XLSX.WorkBook = XLSX.read(data);
      
      let jsonData: string = ""
      if (conversionMode === "simple") {
        // Convert workbook to JSON (each sheet as an array of rows)
        const sheetData: ExcelSheetData = workbook.SheetNames.reduce((acc: ExcelSheetData, sheetName: string) => {
          const sheet: XLSX.WorkSheet = workbook.Sheets[sheetName];
          const jsonData: Array<Record<string, any>> = XLSX.utils.sheet_to_json(sheet);
          acc[sheetName] = jsonData;
          return acc;
        }, {});
        jsonData = JSON.stringify(sheetData, null, 2);
      } else if (conversionMode === "full") {
        jsonData = JSON.stringify(workbook, null, 2)
      }
      
      // Write JSON data to the output file
      fs.writeFileSync(outputJsonPath, jsonData, 'utf-8');
      console.log('\n', `Excel data has been successfully converted and saved to: ${outputJsonPath}`, '\n');
    } catch (err) {
      console.error('Error reading file:', err);
    }
  
  }


  async convertJSONToExcel(conversionMode: string, inputJsonPath: string, outputExcelPath: string): Promise<void> {

    try {
      // Read the JSON file
      const jsonData = fs.readFileSync(inputJsonPath, 'utf-8');
      
      let workbook: XLSX.WorkBook;
      // Parse the JSON data
      const parsedJSONData: any = JSON.parse(jsonData);
      
      if (conversionMode === "simple") {
        
        if (parsedJSONData.hasOwnProperty("bookType")) {
          console.log('\n'+'Error: This JSON data saved in other mode. Use "full" mode to convert data.'+'\n');
          return;
        }
        // Create a new workbook
        workbook = XLSX.utils.book_new();
      
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
      } else if (conversionMode === "full") {
        workbook = parsedJSONData;
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

const conversionModes = ["simple", "full"]

const conversionDirection: string = process.argv[2];
const conversionMode: string = process.argv[3];
if (!conversionModes.includes(conversionMode)) {
  console.log('Error: Exists only "simple" and "full" conversion mode, check command.');
} else {
  const sourcePath: string = process.argv[4];
  const destinationPath: string = process.argv[5];
  
  const excel_convertor = new ExcelConvertor()
  Promise.resolve(excel_convertor.chooseConvertor(conversionDirection, conversionMode, sourcePath, destinationPath));
}

