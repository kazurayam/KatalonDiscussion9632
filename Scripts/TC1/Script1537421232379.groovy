import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import org.apache.poi.ss.format.CellDateFormatter
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellValue
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import com.kms.katalon.core.configuration.RunConfiguration

/**
 * This script will open ./Data Files/WORKDAY_EXAMPLE.xlsx using Apache POI,
 * will print the contents of the Sheet1
 *
 */

Path storageDir = Paths.get(RunConfiguration.getProjectDir()).resolve('Data Files')
Path excelFile = storageDir.resolve('WORKDAY_EXAMPLE.xlsx')

XSSFWorkbook wb

if (Files.exists(excelFile)) {
	FileInputStream fis = new FileInputStream(excelFile.toFile())
	wb = new XSSFWorkbook(fis)
} else {
	throw new IOException("${excelFile.toString()} is not found")
}

XSSFSheet sheet = wb.getSheet('Sheet1')
assert sheet != null

StringBuilder sb = new StringBuilder()
for (int i = 0; i < 10; i++) {
	XSSFRow row = sheet.getRow(i)
	if (row != null) {
		XSSFCell cell0 = row.getCell(0)
		XSSFCell cell1 = row.getCell(1)
		if (cell0 != null && cell1 != null) {
			int type0 = cell0.getCellType()
			int type1 = cell1.getCellType()
			//sb.append("type0=${type0}, type1=${type1}\n")
			// 0: Cell.CELL_TYPE_NUMERIC
			// 1: Cell.CELL_TYPE_STRING
			// 2: Cell.CELL_TYPE_FORMULA
			// 3: Cell.CELL_TYPE_BLANK
			// 4: Cell.CELL_TYPE_BOOLEAN
			// 5: Cell.CELL_TYPE_ERROR
			
			//Date date = cell1.getDateCellValue()
			//SimpleDateFormat sdf = new SimpleDateFormat('dd-MMM-yyyy')
			//sb.append("${cell0},${getCachedFormulaResult(cell1)}\n")
		
			sb.append("${cell0}, ${internallyGetCellText(wb, sheet, i, 1)}\n")
		}
		
	}
}
println "${sb.toString()}"



/**
 * https://stackoverflow.com/questions/7608511/java-poi-how-to-read-excel-cell-value-and-not-the-formula-computing-it
 */
String getCachedFormulaResult(XSSFCell cell) {
	if(cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
		//System.out.println("Formula is " + cell.getCellFormula());
		switch(cell.getCachedFormulaResultType()) {
			case Cell.CELL_TYPE_NUMERIC:
				return "${cell.getNumericCellValue()}"
				break;
			case Cell.CELL_TYPE_STRING:
				return cell.getRichStringCellValue()
				break;
		}
	 }
}

/**
 * copied from com.kms.katalon.core.testdata.reader.SheetPOI call
 * 
 * @param sheet
 * @param col
 * @param row
 * @return
 */
private String internallyGetCellText(Workbook workbook, Sheet sheet, int row, int col) {
	Row curRow = sheet.getRow(row);
	if (curRow == null) {
		return "";
	}
	Cell curCell = curRow.getCell(col);
	if (curCell == null) {
		return "";
	}
	switch (curCell.getCellType()) {
	case 1: // string value
		return curCell.getRichStringCellValue().getString();
	case 0:
		DataFormatter formatter = new DataFormatter(Locale.getDefault());
		return formatter.formatRawCellContents(curCell.getNumericCellValue(), -1,
			getFormatString(curCell.getCellStyle().getDataFormatString()));
	case 4:
		return Boolean.toString(curCell.getBooleanCellValue());
	case 2: // formula
		FormulaEvaluator formulaEval = null;
		try {
			formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
			CellValue cellVal = formulaEval.evaluate(curCell);
			switch (cellVal.getCellType()) {
			case 3:
				return "";
			case 1:
				return cellVal.getStringValue();
			case 0:
				DataFormatter formatter = new DataFormatter(Locale.getDefault());
				return formatter.formatRawCellContents(cellVal.getNumberValue(), -1,
					getFormatString(curCell.getCellStyle().getDataFormatString()));
			}
			return cellVal.formatAsString();
		}
		catch (Exception localException1) {
			try {
				if (DateUtil.isCellDateFormatted(curCell)) {
					String cellFormatString = curCell.getCellStyle().getDataFormatString();
					return new CellDateFormatter(cellFormatString).simpleFormat(curCell.getDateCellValue());
				}
				DataFormatter formatter = new DataFormatter(Locale.getDefault());		
				return formatter.formatRawCellContents(curCell.getNumericCellValue(), -1,
					getFormatString(curCell.getCellStyle().getDataFormatString()));
			}
			catch (Exception localException2) {
				return curCell.getStringCellValue();
			}
		}
	}
	return curCell.getStringCellValue();
}

protected String getFormatString(String rawFormatString) {
	if ((rawFormatString == null) || (rawFormatString.isEmpty())) {
		return rawFormatString;
	}
	return rawFormatString.replace("_(*", "_(\"\"*");
}