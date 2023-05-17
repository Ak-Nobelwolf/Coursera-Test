// Task 1: Build a function-based console log message generator
function consoleStyler(color, background, fontSize, txt) {
    var message = "%c" + txt;
    var style = `color: ${color};`
    style += `background: ${background};`
    style += `font-size: ${fontSize};`
    console.log(message, style);
  }
  
  // Task 2: Build another console log message generator
  function celebrateStyler(reason) {
    var fontStyle = "color: tomato; font-size: 50px";
    if (reason == "birthday") {
        console.log(`%cHappy birthday`, fontStyle);
    } else if (reason == "champions") {
        console.log(`%cCongrats on the title!`, fontStyle);
    } else {
        console.log(message, style);
    }
  
  // Task 3: Run both the consoleStyler and the celebrateStyler functions
    consoleStyler('#1d5c63', '#ede6db', '40px', 'Congrats!');
    consoleStyler('birthday');
  
  // Task 4: Insert a congratulatory and custom message
  function styleAndCelebrate(color, background, fontSize, txt, reason) {
    consoleStyler(color, background, fontSize, txt);  
    celebrateStyler("reason");
  
  }
  // Call styleAndCelebrate
    styleAndCelebrate('ef7c8e', 'fae8e0', '30px', 'You made it!', 'champions');
  }

  /*
/*
This Java program processes an Excel file by splitting its data into multiple Excel files based on the data in the first column. 
The program reads an input Excel file, extracts the column headings from the first row, and then processes each row in the sheet. 
Each row is checked to determine if it should be included in a new Excel file or not. 
If the first column of a row contains a non-empty string, the program creates a new Excel file (if it doesn't exist yet) and adds the row to the corresponding sheet in the new Excel file.
The program also clones the cell style of each cell in the input file to the corresponding cell in the new Excel files.

The program uses the Apache POI library to read and write Excel files. The main() method takes the path of the input Excel file as an argument, and then performs the following steps:
1. Reads the input Excel file using the WorkbookFactory.create() method.
2. Extracts the column headings from the first row and stores them in a map, with the column names as keys and the column indices as values.
3. Processes each row in the sheet, and creates a new Excel file if necessary.
4. Clones the cell style of each cell in the input file to the corresponding cell in the new Excel files.
5. Saves the new Excel files to the output folder.
6. Closes the input Excel file.

Note that the program assumes that the data in the first column is used to determine the file name for each row, and that the file name does not contain any illegal characters. 
If the first column contains duplicate values, the program will overwrite the corresponding files.
 */

//V.0.01
/* 
package pkg;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProcessor {

	public static void main(String[] args) {
		// path of the input Excel file
		String filePath = "C:\\Users\\akshay.tadakod\\Documents\\output\\Excel File\\History_Sample.xlsx";

		try {
			// read the input Excel file
			Workbook workbook = WorkbookFactory.create(new File(filePath));
			// assuming the data is in the first sheet
			Sheet sheet = workbook.getSheetAt(0);

			// get the column headings from the first row
			Row headerRow = sheet.getRow(0);
			Map<String, Integer> columnIndices = new HashMap<>();
			for (int i = 0; i < headerRow.getLastCellNum(); i++) {
				Cell cell = headerRow.getCell(i);
				if (cell != null && cell.getCellType() == CellType.STRING) {
					String columnName = cell.getStringCellValue().trim();
					columnIndices.put(columnName, i);
				}
			}

			// split the data into multiple Excel files based on the data in the first
			// column
			Map<String, Workbook> workbooks = new HashMap<>();
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row dataRow = sheet.getRow(i);
				if (dataRow == null) {
					continue; // skip empty rows
				}

				Cell firstColumnCell = dataRow.getCell(0);
				if (firstColumnCell == null || firstColumnCell.getCellType() != CellType.STRING) {
					continue; // skip rows with invalid data in the first column
				}

				String fileName = firstColumnCell.getStringCellValue().trim();
				if (fileName.length() == 0) {
					continue; // skip rows with empty file name
				}

				Workbook newWorkbook = workbooks.get(fileName);
				if (newWorkbook == null) {
					// create a new Excel file with the same column headings as the input file
					newWorkbook = WorkbookFactory.create(true);
					Sheet newSheet = newWorkbook.createSheet();
					Row newHeaderRow = newSheet.createRow(0);
					for (String columnName : columnIndices.keySet()) {
						int columnIndex = columnIndices.get(columnName);
						Cell newCell = newHeaderRow.createCell(columnIndex);
						newCell.setCellValue(columnName);
					}
					workbooks.put(fileName, newWorkbook);
				}

				// add the current row to the new Excel file
				int rowIndex = newWorkbook.getSheetAt(0).getLastRowNum() + 1;
				Row newDataRow = newWorkbook.getSheetAt(0).createRow(rowIndex);
				for (int j = 0; j < dataRow.getLastCellNum(); j++) {
					Cell dataCell = dataRow.getCell(j);
					if (dataCell != null) {
						Cell newDataCell = newDataRow.createCell(j);
						newDataCell.setCellStyle(newWorkbook.createCellStyle()); // create a new style for the target
						// workbook
						newDataCell.getCellStyle().cloneStyleFrom(dataCell.getCellStyle()); // copy properties from the
						// source style
						switch (dataCell.getCellType()) {
						case STRING:
							newDataCell.setCellValue(dataCell.getStringCellValue());
							break;
						case NUMERIC:
							newDataCell.setCellValue(dataCell.getNumericCellValue());
							break;
						case BOOLEAN:
							newDataCell.setCellValue(dataCell.getBooleanCellValue());
							break;
						case FORMULA:
							newDataCell.setCellFormula(dataCell.getCellFormula());
							break;
						case BLANK:
							// do nothing
							break;
						case ERROR:
							newDataCell.setCellErrorValue(dataCell.getErrorCellValue());
							break;
						default:
							// do nothing
							break;
						}
					}
				}
			}

			// save the new Excel files
			for (String fileName : workbooks.keySet()) {
				// Path for the output/extracted files
				String outputFilePath = "C:\\Users\\akshay.tadakod\\Documents\\output\\Excel File\\" + fileName + ".xlsx";
				FileOutputStream outputStream = new FileOutputStream(outputFilePath);
				Workbook newWorkbook = workbooks.get(fileName);
				newWorkbook.write(outputStream);
				newWorkbook.close();
			}

			// close the input Excel file
			workbook.close();
			System.out.println("Excel Files Extracted Successfully!");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
 */


//V.0.02 - Modified the code to Replace new lines, extra spaces, and tab spaces with commas
/*
package pkg;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProcessor {

	public static void main(String[] args) {
		// path of the input Excel file
		String filePath = "C:\\Users\\akshay.tadakod\\Documents\\output\\Excel File\\History_Sample.xlsx";

		try {
			// read the input Excel file
			Workbook workbook = WorkbookFactory.create(new File(filePath));
			// assuming the data is in the first sheet
			Sheet sheet = workbook.getSheetAt(0);

			// get the column headings from the first row
			Row headerRow = sheet.getRow(0);
			Map<String, Integer> columnIndices = new HashMap<>();
			for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			    Cell cell = headerRow.getCell(i);
			    if (cell != null && cell.getCellType() == CellType.STRING) {
			        String columnName = cell.getStringCellValue().trim();
			        columnIndices.put(columnName, i);
			    }
			}


			// split the data into multiple Excel files based on the data in the first column
			Map<String, Workbook> workbooks = new HashMap<>();
			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row dataRow = sheet.getRow(i);
				if (dataRow == null) {
					continue; // skip empty rows
				}

				Cell firstColumnCell = dataRow.getCell(0);
				if (firstColumnCell == null || firstColumnCell.getCellType() != CellType.STRING) {
					continue; // skip rows with invalid data in the first column
				}

				String fileName = firstColumnCell.getStringCellValue().trim();
				if (fileName.length() == 0) {
					continue; // skip rows with empty file name
				}

				Workbook newWorkbook = workbooks.get(fileName);
				if (newWorkbook == null) {
					// create a new Excel file with the same column headings as the input file
					newWorkbook = WorkbookFactory.create(true);
					Sheet newSheet = newWorkbook.createSheet();
					Row newHeaderRow = newSheet.createRow(0);
					for (String columnName : columnIndices.keySet()) {
						int columnIndex = columnIndices.get(columnName);
						Cell newCell = newHeaderRow.createCell(columnIndex);
						newCell.setCellValue(columnName);
					}
					workbooks.put(fileName, newWorkbook);
				}

				// add the current row to the new Excel file
				int rowIndex = newWorkbook.getSheetAt(0).getLastRowNum() + 1;
				Row newDataRow = newWorkbook.getSheetAt(0).createRow(rowIndex);
				for (int j = 0; j < dataRow.getLastCellNum(); j++) {
				    Cell dataCell = dataRow.getCell(j);
				    if (dataCell != null) {
				        Cell newDataCell = newDataRow.createCell(j);
				        newDataCell.setCellStyle(newWorkbook.createCellStyle()); // create a new style for the target workbook
				        newDataCell.getCellStyle().cloneStyleFrom(dataCell.getCellStyle()); // copy properties from the source style
				        if (j == columnIndices.get("Details")) {
				            // Handle the "Details" column
				            String detailsValue = dataCell.getStringCellValue().trim();
				            detailsValue = detailsValue.replaceAll("[\\r\\n\\t]+", ","); // Replace new lines, extra spaces, and tab spaces with commas
				            newDataCell.setCellValue(detailsValue);
				        } else {
				            // Copy other cell values as it is
				            switch (dataCell.getCellType()) {
				                case STRING:
				                    newDataCell.setCellValue(dataCell.getStringCellValue());
				                    break;
				                case NUMERIC:
				                    newDataCell.setCellValue(dataCell.getNumericCellValue());
				                    break;
				                case BOOLEAN:
				                    newDataCell.setCellValue(dataCell.getBooleanCellValue());
				                    break;
				                case FORMULA:
				                    newDataCell.setCellFormula(dataCell.getCellFormula());
				                    break;
				                case BLANK:
				                    // do nothing
				                    break;
				                case ERROR:
				                    newDataCell.setCellErrorValue(dataCell.getErrorCellValue());
				                    break;
				                default:
				                    // do nothing
				                    break;
				            }
				        }
				    }
				}
			}


			// save the new Excel files
			for (String fileName : workbooks.keySet()) {
				// Path for the output/extracted files
				String outputFilePath = "C:\\Users\\akshay.tadakod\\Documents\\output\\Excel File\\" + fileName + ".xlsx";
				FileOutputStream outputStream = new FileOutputStream(outputFilePath);
				Workbook newWorkbook = workbooks.get(fileName);
				newWorkbook.write(outputStream);
				newWorkbook.close();
			}

			// close the input Excel file
			workbook.close();
			System.out.println("Excel Files Extracted Successfully!");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
 */

//V.0.03.00 - 
//1. Modified the code to have only 5000 rows in a sheet and continue remaining in a next sheet 
//2. Modified the code to convert the date and time to YYYY/MM/DD HH24:MI:SS format and time zone to device local time 

/*
Below code processes an input Excel file and splits its data into multiple Excel files based on the values in the first column. It uses the Apache POI library for reading and writing Excel files.

Here's a breakdown of the code:
1. The program starts by importing the necessary packages and classes.
2. The ExcelProcessor class is defined, which contains the main method for executing the program.
3. Several constants are defined, including file paths, column names, and maximum rows per split.
4. The main method begins by reading the input Excel file using the WorkbookFactory class from Apache POI.
5. The first sheet of the workbook is obtained, assuming that's where the data resides.
6. The column headings are extracted from the first row to create a mapping of column names to column indices.
7. The data rows are iterated starting from the second row. Empty rows and rows with null data in the first column are skipped.
8. The first column cell is checked for its type (string or numeric). The cell value is then trimmed and stored as the file name.
9. If the file name is empty, the row is skipped.
10. The program keeps track of the row count for each file name. When the row count exceeds the maximum rows per split or a new file needs to be created, a new workbook is created and added to the workbooks map.
11. The data row is added to the corresponding workbook based on the file name and row count.
12. After iterating through all the rows, the program saves each workbook as a separate Excel file in the specified output directory.
13. Finally, the input workbook is closed, and a success message is printed.

The code also includes helper methods for extracting column indices, creating new workbooks, adding data rows to workbooks, parsing date and time values, and saving workbooks.
 */

/*
package pkg;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProcessor {
	private static final int FIRST_ROW_INDEX = 0;
	private static final int FIRST_COLUMN_INDEX = 0;
	private static final String DETAILS_COLUMN_NAME = "Details";
	private static final Path FILE_PATH = Paths.get("C:", "Users", "akshay.tadakod", "Documents", "output", "Excel File", "History_Sample.xlsx");
	private static final Path OUTPUT_DIRECTORY = Paths.get("C:", "Users", "akshay.tadakod", "Documents", "output", "Excel File", "Extracted files");

	private static final int MAX_ROWS_PER_SPLIT = 4999;

	private static final String LOCAL_CLIENT_TIME_COLUMN_NAME = "Local Client Time";
	private static final DateTimeFormatter OUTPUT_DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");

	public static void main(String[] args) {
		try {
			// Read the input Excel file
			Workbook workbook = WorkbookFactory.create(FILE_PATH.toFile());
			// Assuming the data is in the first sheet
			Sheet sheet = workbook.getSheetAt(0);

			// Get the column headings from the first row
			Row headerRow = sheet.getRow(FIRST_ROW_INDEX);
			Map<String, Integer> columnIndices = getColumnIndices(headerRow);

			// Split the data into multiple Excel files based on the data in the first column
			Map<String, Workbook> workbooks = new HashMap<>();
			Map<String, Integer> rowCounts = new HashMap<>();

			int rowIndex = 1; // Start from row 1 as row 0 is the header
			int fileIndex = 0; // Start with file index 0

			while (rowIndex <= sheet.getLastRowNum()) {
				Row dataRow = sheet.getRow(rowIndex);
				if (dataRow == null) {
					rowIndex++;
					continue; // Skip empty rows
				}

				Cell firstColumnCell = dataRow.getCell(FIRST_COLUMN_INDEX);
				if (firstColumnCell == null) {
					rowIndex++;
					continue; // Skip rows with null data in the first column
				}

				String fileName;
				if (firstColumnCell.getCellType() == CellType.STRING) {
					fileName = firstColumnCell.getStringCellValue().trim();
				} else if (firstColumnCell.getCellType() == CellType.NUMERIC) {
					fileName = String.valueOf((int) firstColumnCell.getNumericCellValue()).trim();
				} else {
					rowIndex++;
					continue; // Skip rows with unsupported data type in the first column
				}

				if (fileName.isEmpty()) {
					rowIndex++;
					continue; // Skip rows with an empty file name
				}

				int rowCount = rowCounts.getOrDefault(fileName, 0);
				Workbook currentWorkbook = workbooks.get(fileName + "_" + fileIndex);
				if (currentWorkbook == null || rowCount >= MAX_ROWS_PER_SPLIT) {
					fileIndex++;
					rowCount = 0;
					// Create a new workbook
					currentWorkbook = createNewWorkbook(columnIndices);
					workbooks.put(fileName + "_" + fileIndex, currentWorkbook);
				}
				addDataRowToWorkbook(dataRow, currentWorkbook, columnIndices);

				rowCounts.put(fileName, rowCount + 1);
				rowIndex++;
			}
			// Save the new Excel files
			for (Map.Entry<String, Workbook> entry : workbooks.entrySet()) {
				String fileName = entry.getKey();
				Workbook newWorkbook = entry.getValue();
				String filePath = OUTPUT_DIRECTORY.resolve(fileName + ".xlsx").toString();
				saveWorkbook(newWorkbook, filePath);
			}
			// Close the input Excel file
			workbook.close();
			System.out.println("Excel Files Extracted Successfully!.");
		} catch (IOException e) {
			handleIOException(e);
		}
	}

	private static Map<String, Integer> getColumnIndices(Row headerRow) {
		Map<String, Integer> columnIndices = new HashMap<>();
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			Cell cell = headerRow.getCell(i);
			if (cell != null && cell.getCellType() == CellType.STRING) {
				String columnName = cell.getStringCellValue().trim();
				columnIndices.put(columnName, i);
			}
		}
		return columnIndices;
	}

	private static Workbook createNewWorkbook(Map<String, Integer> columnIndices) {
		try {
			Workbook newWorkbook = WorkbookFactory.create(true);
			Sheet newSheet = newWorkbook.createSheet();
			Row newHeaderRow = newSheet.createRow(FIRST_ROW_INDEX);
			for (Map.Entry<String, Integer> entry : columnIndices.entrySet()) {
				String columnName = entry.getKey();
				int columnIndex = entry.getValue();
				Cell newCell = newHeaderRow.createCell(columnIndex);
				newCell.setCellValue(columnName);
			}
			return newWorkbook;
		} catch (IOException e) {
			handleIOException(e);
			return null;
		}
	}

	private static void addDataRowToWorkbook(Row dataRow, Workbook newWorkbook, Map<String, Integer> columnIndices) {
		Sheet newSheet = newWorkbook.getSheetAt(0);
		if (newSheet == null) {
			newSheet = newWorkbook.createSheet();
		}
		int rowIndex = newSheet.getLastRowNum() + 1;
		Row newDataRow = newSheet.createRow(rowIndex);

		Map<Short, CellStyle> cellStyles = new HashMap<>();

		for (int j = 0; j < dataRow.getLastCellNum(); j++) {
			Cell dataCell = dataRow.getCell(j);
			if (dataCell != null) {
				Cell newDataCell = newDataRow.createCell(j);

				// Retrieve or create the cell style
				CellStyle cellStyle = cellStyles.get(dataCell.getCellStyle().getIndex());
				if (cellStyle == null) {
					cellStyle = newWorkbook.createCellStyle();
					cellStyle.cloneStyleFrom(dataCell.getCellStyle());
					cellStyles.put(dataCell.getCellStyle().getIndex(), cellStyle);
				}

				newDataCell.setCellStyle(cellStyle);
				if (j == columnIndices.get(DETAILS_COLUMN_NAME) && j == columnIndices.get(LOCAL_CLIENT_TIME_COLUMN_NAME)) {
					// Handle the columns that meet both conditions

					// Handle the "Details" column
					String detailsValue = dataCell.getStringCellValue().trim();
					detailsValue = detailsValue.replaceAll("[\r\n\t]+", ",");
					newDataCell.setCellValue(detailsValue);

					// Handle the "Local Client Time" column
					String dateTimeValue = dataCell.getStringCellValue().trim();
					LocalDateTime localDateTime = parseDateTime(dateTimeValue);
					if (localDateTime != null) {
						String formattedDateTime = OUTPUT_DATE_TIME_FORMATTER.format(localDateTime);
						newDataCell.setCellValue(formattedDateTime);
					} else {
						newDataCell.setCellValue(dateTimeValue);
					}

				} else if (j == columnIndices.get(DETAILS_COLUMN_NAME)) {
					// Handle the "Details" column
					String detailsValue = dataCell.getStringCellValue().trim();
					detailsValue = detailsValue.replaceAll("[\r\n\t]+", ",");
					newDataCell.setCellValue(detailsValue);

				} else if (j == columnIndices.get(LOCAL_CLIENT_TIME_COLUMN_NAME)) {
					// Handle the "Local Client Time" column
					String dateTimeValue = dataCell.getStringCellValue().trim();
					LocalDateTime localDateTime = parseDateTime(dateTimeValue);
					if (localDateTime != null) {
						String formattedDateTime = OUTPUT_DATE_TIME_FORMATTER.format(localDateTime);
						newDataCell.setCellValue(formattedDateTime);
					} else {
						newDataCell.setCellValue(dateTimeValue);
					}
				}

//								if (j == columnIndices.get(DETAILS_COLUMN_NAME)) {
//									// Handle the "Details" column
//									String detailsValue = dataCell.getStringCellValue().trim();
//									detailsValue = detailsValue.replaceAll("[\r\n\t]+", ",");
//									newDataCell.setCellValue(detailsValue);
//								} if (j == columnIndices.get(LOCAL_CLIENT_TIME_COLUMN_NAME)) {
//									// Handle the "Local Client Time" column
//									String dateTimeValue = dataCell.getStringCellValue().trim();
//									LocalDateTime localDateTime = parseDateTime(dateTimeValue);
//									if (localDateTime != null) {
//										String formattedDateTime = OUTPUT_DATE_TIME_FORMATTER.format(localDateTime);
//										newDataCell.setCellValue(formattedDateTime);
//									} else {
//										newDataCell.setCellValue(dateTimeValue);
//									}
//								} 
				else {
					// Copy other cell values as it is
					switch (dataCell.getCellType()) {
					case STRING:
						newDataCell.setCellValue(dataCell.getStringCellValue());
						break;
					case NUMERIC:
						if (DateUtil.isCellDateFormatted(dataCell)) {
							newDataCell.setCellValue(dataCell.getDateCellValue());
						} else {
							newDataCell.setCellValue(dataCell.getNumericCellValue());
						}
						break;
					case BOOLEAN:
						newDataCell.setCellValue(dataCell.getBooleanCellValue());
						break;
					case FORMULA:
						newDataCell.setCellFormula(dataCell.getCellFormula());
						break;
					default:
						newDataCell.setCellValue(dataCell.getStringCellValue());
						break;
					}
				}
			}
		}
	}
	private static LocalDateTime parseDateTime(String dateTimeValue) {
		SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss aa");
		try {
			Date date = dateFormat.parse(dateTimeValue);
			return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
		} catch (ParseException e) {
			handleParseException(e);
			return null;
		}
	}

	private static void saveWorkbook(Workbook workbook, String filePath) {
		try (OutputStream outputStream = new FileOutputStream(filePath)) {
			workbook.write(outputStream);
			System.out.println("File saved: " + filePath);
		} catch (IOException e) {
			handleIOException(e);
		}
	}

	private static void handleIOException(IOException e) {
		System.out.println("An IO exception occurred: " + e.getMessage());
	}

	private static void handleParseException(ParseException e) {
		System.out.println("Error parsing date/time: " + e.getMessage());
	}
}
 */

//V.0.03.01 - 
//1. Modified the code to have only 5000 rows in a sheet and continue remaining in a next sheet 
//2. Modified the code to convert the date and time to YYYY/MM/DD HH24:MI:SS format and time zone to device local time 
/*
package pkg;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProcessor {
	private static final int FIRST_ROW_INDEX = 0;
	private static final int FIRST_COLUMN_INDEX = 0;
	
	//Mention file path for input excel file which is to be extracted. Instead of // use , and " " to specify the folder names
	private static final Path FILE_PATH = Paths.get("C:", "Users", "akshay.tadakod", "Documents", "output", "Excel File", "Date_Test 2.xlsx");
	
	//Mention folder/directory  path to store extracted files. Instead of // use , and " " to specify the folder names
	private static final Path OUTPUT_DIRECTORY = Paths.get("C:", "Users", "akshay.tadakod", "Documents", "output", "Excel File", "Extracted files");

	//Limits the file to have only 4999 records per page excluding the header, so including header row its 5000 rows per file.
	private static final int MAX_ROWS_PER_SPLIT = 4999; 
	
	//Change with column name in which you have new lines, extra spaces, and tab spaces to replace with commas
	private static final String DETAILS_COLUMN_NAME = "DETAILS"; 
	
	//Change with column name which you want to convert to below mentioned(i.e., "yyyy/MM/dd HH:mm:ss") date time format
	private static final String TIMESTAMP_COLUMN_NAME = "TIMESTAMP"; 
	
	//Read line 805
	private static final DateTimeFormatter OUTPUT_DATE_TIME_FORMATTER = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");

	public static void main(String[] args) {
		try {
			// Read the input Excel file
			Workbook workbook = WorkbookFactory.create(FILE_PATH.toFile());
			// Assuming the data is in the first sheet
			Sheet sheet = workbook.getSheetAt(0);

			// Get the column headings from the first row
			Row headerRow = sheet.getRow(FIRST_ROW_INDEX);
			Map<String, Integer> columnIndices = getColumnIndices(headerRow);

			// Split the data into multiple Excel files based on the data in the first column
			Map<String, Workbook> workbooks = new HashMap<>();
			Map<String, Integer> rowCounts = new HashMap<>();

			int rowIndex = 1; // Start from row 1 as row 0 is the header
			int fileIndex = 0; // Start with file index 0

			while (rowIndex <= sheet.getLastRowNum()) {
				Row dataRow = sheet.getRow(rowIndex);
				if (dataRow == null) {
					rowIndex++;
					continue; // Skip empty rows
				}

				Cell firstColumnCell = dataRow.getCell(FIRST_COLUMN_INDEX);
				if (firstColumnCell == null) {
					rowIndex++;
					continue; // Skip rows with null data in the first column
				}

				String fileName;
				if (firstColumnCell.getCellType() == CellType.STRING) {
					fileName = firstColumnCell.getStringCellValue().trim();
				} else if (firstColumnCell.getCellType() == CellType.NUMERIC) {
					fileName = String.valueOf((int) firstColumnCell.getNumericCellValue()).trim();
				} else {
					rowIndex++;
					continue; // Skip rows with unsupported data type in the first column
				}

				if (fileName.isEmpty()) {
					rowIndex++;
					continue; // Skip rows with an empty file name
				}

				int rowCount = rowCounts.getOrDefault(fileName, 0);
				Workbook currentWorkbook = workbooks.get(fileName + "_" + fileIndex);
				if (currentWorkbook == null || rowCount >= MAX_ROWS_PER_SPLIT) {
					fileIndex++;
					rowCount = 0;
					// Create a new workbook
					currentWorkbook = createNewWorkbook(columnIndices);
					workbooks.put(fileName + "_" + fileIndex, currentWorkbook);
				}
				addDataRowToWorkbook(dataRow, currentWorkbook, columnIndices);

				rowCounts.put(fileName, rowCount + 1);
				rowIndex++;
			}
			// Save the new Excel files
			for (Map.Entry<String, Workbook> entry : workbooks.entrySet()) {
				String fileName = entry.getKey();
				Workbook newWorkbook = entry.getValue();
				String filePath = OUTPUT_DIRECTORY.resolve(fileName + ".xlsx").toString();
				saveWorkbook(newWorkbook, filePath);
			}
			// Close the input Excel file
			workbook.close();
			System.out.println("Excel Files Extracted Successfully!.");
		} catch (IOException e) {
			handleIOException(e);
		}
	}

	private static Map<String, Integer> getColumnIndices(Row headerRow) {
		Map<String, Integer> columnIndices = new HashMap<>();
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			Cell cell = headerRow.getCell(i);
			if (cell != null && cell.getCellType() == CellType.STRING) {
				String columnName = cell.getStringCellValue().trim();
				columnIndices.put(columnName, i);
			}
		}
		return columnIndices;
	}

	private static Workbook createNewWorkbook(Map<String, Integer> columnIndices) {
		try {
			Workbook newWorkbook = WorkbookFactory.create(true);
			Sheet newSheet = newWorkbook.createSheet();
			Row newHeaderRow = newSheet.createRow(FIRST_ROW_INDEX);
			for (Map.Entry<String, Integer> entry : columnIndices.entrySet()) {
				String columnName = entry.getKey();
				int columnIndex = entry.getValue();
				Cell newCell = newHeaderRow.createCell(columnIndex);
				newCell.setCellValue(columnName);
			}
			return newWorkbook;
		} catch (IOException e) {
			handleIOException(e);
			return null;
		}
	}

	private static void addDataRowToWorkbook(Row dataRow, Workbook newWorkbook, Map<String, Integer> columnIndices) {
		Sheet newSheet = newWorkbook.getSheetAt(0);
		if (newSheet == null) {
			newSheet = newWorkbook.createSheet();
		}
		int rowIndex = newSheet.getLastRowNum() + 1;
		Row newDataRow = newSheet.createRow(rowIndex);

		Map<Short, CellStyle> cellStyles = new HashMap<>();

		for (int j = 0; j < dataRow.getLastCellNum(); j++) {
			Cell dataCell = dataRow.getCell(j);
			if (dataCell != null) {
				Cell newDataCell = newDataRow.createCell(j);

				// Retrieve or create the cell style
				CellStyle cellStyle = cellStyles.get(dataCell.getCellStyle().getIndex());
				if (cellStyle == null) {
					cellStyle = newWorkbook.createCellStyle();
					cellStyle.cloneStyleFrom(dataCell.getCellStyle());
					cellStyles.put(dataCell.getCellStyle().getIndex(), cellStyle);
				}

				newDataCell.setCellStyle(cellStyle);
				if (j == columnIndices.get(DETAILS_COLUMN_NAME)) {
					// Handle the "Details" column
					String detailsValue = dataCell.getStringCellValue().trim();
					detailsValue = detailsValue.replaceAll("[\r\n\t]+", ",");
					newDataCell.setCellValue(detailsValue);
				} 
				else if (j == columnIndices.get(TIMESTAMP_COLUMN_NAME)) {
					// Handle the "Local Client Time" column
					DataFormatter formatter = new DataFormatter();
					String formattedvalue = formatter.formatCellValue(dataCell);	
					LocalDateTime localDateTime = parseDateTime(formattedvalue);
					if (localDateTime != null) {
						String formattedDateTime = OUTPUT_DATE_TIME_FORMATTER.format(localDateTime);
						newDataCell.setCellValue(formattedDateTime);
					} else {
						newDataCell.setCellValue(formattedvalue);
					}
				}
				else {
					// Copy other cell values as it is
					switch (dataCell.getCellType()) {
					case STRING:
						newDataCell.setCellValue(dataCell.getStringCellValue());
						break;
					case NUMERIC:
						if (DateUtil.isCellDateFormatted(dataCell)) {
							newDataCell.setCellValue(dataCell.getDateCellValue());
						} else {
							newDataCell.setCellValue(dataCell.getNumericCellValue());
						}
						break;
					case BOOLEAN:
						newDataCell.setCellValue(dataCell.getBooleanCellValue());
						break;
					case FORMULA:
						newDataCell.setCellFormula(dataCell.getCellFormula());
						break;
					default:
						newDataCell.setCellValue(dataCell.getStringCellValue());
						break;
					}
				}
			}
		}
	}
	private static LocalDateTime parseDateTime(String dateTimeValue) {
		//change this format accordingly to excel file which you want to convert to
//		SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy hh:mm:ss aa"); 
//		SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy hh:mm aa");
		SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
		String[] dateFormats = {
				
		};

		try {
			Date date = dateFormat.parse(dateTimeValue);
			return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
		} catch (ParseException e) {
			handleParseException(e);
			return null;
		}
	}

	private static void saveWorkbook(Workbook workbook, String filePath) {
		try (OutputStream outputStream = new FileOutputStream(filePath)) {
			workbook.write(outputStream);
			System.out.println("File saved: " + filePath);
		} catch (IOException e) {
			handleIOException(e);
		}
	}

	private static void handleIOException(IOException e) {
		System.out.println("An IO exception occurred: " + e.getMessage());
	}

	private static void handleParseException(ParseException e) {
		System.out.println("Error parsing date/time: " + e.getMessage());
	}
}
  */