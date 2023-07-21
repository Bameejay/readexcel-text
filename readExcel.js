const Excel = require('exceljs');
const fs = require('fs');
const path = require('path');

const readExcelAndGenerateTextFile = async () => {
  const excelFilePath = path.join(__dirname, 'readexceltotxt.xlsx'); // Replace 'your-file.xlsx' with the actual filename and path

  try {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(excelFilePath);

    let textData = '';

    const worksheet = workbook.getWorksheet('Sheet1'); // Replace 'Sheet1' with the name of your sheet

    worksheet.eachRow((row, rowNumber) => {
      const cell1Value = row.getCell(1).value;
      const cell2Value = row.getCell(2).value;
      // Add more columns if needed

      textData += `Row ${rowNumber}: ${cell1Value}, ${cell2Value}\n`;
    });

    const textFilePath = path.join("C:\\Users\\JOHNSON FRANCIS\\Documents\\vbs testing\\output.txt"); // Replace 'output.txt' with the desired filename and path for your text file

    fs.writeFile(textFilePath, textData, (err) => {
      if (err) {
        console.error(`Error writing to text file: ${err.message}`);
        return;
      }
      console.log(`Text file generated successfully at: ${textFilePath}`);
    });

  } catch (error) {
    console.error(`Error reading Excel file: ${error.message}`);
    return;
  }
};

readExcelAndGenerateTextFile();
