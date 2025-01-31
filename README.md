# Excel convertor to JSON file.


## Install project

```bash
npm install
```

### Check versions
```bash
node -v
npm -v
```
If no one commands need install

## Compile
For compile use this command
```bash
npm run dev
```
It's command recompile main.js file every time if added any changes.

This command compile only once and create main.js.
```bash
npm run build
```


## Run

```bash
node main.js FolderPaths/ExceFile.xlsx
```
Result JSON file will be saved to current folder with this Excel file.

Run with output folder:
```bash
node main.js FolderPaths/ExceFile.xlsx OutputFolderPath
```

If path have spaces need use quotes.
```bash
node main.js "FolderPaths/ExceFile.xlsx" "OutputFolderPath"
```