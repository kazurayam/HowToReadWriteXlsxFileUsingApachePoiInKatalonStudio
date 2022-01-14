/**
 * A sample Test Case script in Katalon Studio that reads and writes an Excel file using Apache POI
 *
 * The original is
 * https://vektorwebsolutions.com/how-to-read-write-xlsx-file-in-java-apach-poi-example/
 */
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/*
 * Reading ./Employee.xlsx file
 */
File sourceExcel = new File("./Employee.xlsx");
FileInputStream fis = new FileInputStream(sourceExcel);

// Open the .xlsx file and construct a workbook object
XSSFWorkbook book = new XSSFWorkbook(fis)

// Get the top sheet out of the workbook 
XSSFSheet sheet = book.getSheetAt(0)

// Iterating over rows in the sheet
for (Row row in sheet) {
	// Iterating over each column of the row
	for (Cell cell in row) {
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				print(cell.getStringCellValue() + "\t")
				break
			case Cell.CELL_TYPE_NUMERIC:
				print(cell.getNumericCellValue() + "\t")
				break
			case Cell.CELL_TYPE_BOOLEAN:
				print(cell.getBooleanCellValue() + "\t")
				break
			default:
				break
		}
	}
	print "\n"
}

	
// writing data into the sheet
Map<String, Object[]> newData = new HashMap<>()
newData.put("7", [ 7d, "Sonya", "75K", "SALES", "Rupert"])
newData.put("8", [ 8d, "Kris", "85K", "SALES", "Rupert" ])
newData.put("9", [ 9d, "Dave", "90K", "SALES", "Rupert" ])

Set newRows = newData.keySet()
//int rownum = sheet.getLastRowNum()
int rownum = findLastRowNum(sheet)
for (String key in newRows) {
	Row row = sheet.createRow(rownum++)
	Object[] objArr = newData.get(key)
	int cellnum = 0
	for (Object obj : objArr) {
		Cell cell = row.createCell(cellnum++)
		if (obj instanceof String) {
			cell.setCellValue((String) obj)
		} else if (obj instanceof Boolean) {
			cell.setCellValue((Boolean) obj)
		} else if (obj instanceof Date) {
			cell.setCellValue((Date) obj)
		} else if (obj instanceof Double) {
			cell.setCellValue((Double) obj)
		}
	}
}

/*
 * Writing into the ./Employee.out.xlsx file
 */
File targetExcel = new File("./Employee.out.xlsx");
FileOutputStream os = new FileOutputStream(targetExcel)
book.write(os)
println "Writing into ${targetExcel}";

// Close workbook, OutputStream and Excel file to prevent leak
os.close();
book.close();
fis.close();


/**
 * find the index of the last row with a valid numeric value in the 1st column ("ID")
 */
int findLastRowNum(Sheet sheet) {
	int index = 0
	int rx = 0
	for (Row row in sheet) {
		rx += 1
		Cell cell = row.getCell(0)
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			index = rx		
		}
	}
	return index
}