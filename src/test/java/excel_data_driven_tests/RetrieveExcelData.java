package excel_data_driven_tests;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RetrieveExcelData {

	// Method to return data values from specified worksheet, first row header and test case
	public ArrayList<String> getData(String testSheet, String testCase, String testData) throws IOException {
		// TODO Auto-generated method stub

		// Create ArrayList to store data
		ArrayList<String> dataList = new ArrayList<String>();
		
		// Create an object that can read the Excel file
		FileInputStream fileIS = new FileInputStream("C:\\Users\\Khalid\\Documents\\Documents\\Courses\\Selenium\\Section 26\\Data.xlsx");
		
		// Create an object that can access the Excel file
		XSSFWorkbook fileWB = new XSSFWorkbook(fileIS);
		
		// Retrieve Test Data worksheet
		XSSFSheet testDataSheet = fileWB.getSheet(testSheet);
		
		// Retrieve first row
		
		// Create iterator of all rows in testDataSheet
		Iterator<Row> rows = testDataSheet.rowIterator();
		
		// Retrieve first row in testDataSheet
		Row firstRow = rows.next();
		
		// Identify test case cell in firstRow
		
		// Create iterator of all cells in firstRow
		Iterator<Cell> firstRowCells = firstRow.cellIterator();
		
		// Iterate through firstRowCells until test case cell found
		
		// While firstRowCells has another cell
		while (firstRowCells.hasNext())
		{
			// Retrieve next cell in firstRowCells
			Cell firstRowCell = firstRowCells.next();
						
			// If value of cell is test case
			if (stringValue(firstRowCell).equalsIgnoreCase(testCase))
			{
				// Store column index of test case column
				int testCaseColumnIndex = firstRowCell.getColumnIndex();
				
				// Identify test data row
				
				// While worksheet has another data row
				while (rows.hasNext())
				{
					// Retrieve next row
					Row nextRow = rows.next();
					
					// Retrieve value of test case in row
					String testDataCase = stringValue(nextRow.getCell(testCaseColumnIndex));
					
					// If test case is for passed testData
					if (testDataCase.equalsIgnoreCase(testData))
					{
						// Create iterator of all cells in test data row
						Iterator<Cell> nextRowCells = nextRow.cellIterator();
						
						// While data exists in test data row
						while (nextRowCells.hasNext())
						{
							// Add data to dataList array list
							dataList.add(stringValue(nextRowCells.next()));
						}
						
						// For each cell before actual test data
						for (int i=0; i<=testCaseColumnIndex; i++)
						{
							// Remove data from dataList array list
							dataList.remove(0);
						}
					}
				}
			}
		}
		// Close link to Excel file
		fileWB.close();
		
		// Return data list
		return dataList;
	}
	
	// Method to convert non-String values in Excel file to appropriate Strings
	public static String stringValue(Cell c)
	{
		// Retrieve cell type of passed cell
		CellType cType = c.getCellType();
		
		// Create String to be returned
		String stringCValue = "";
		
		// If cell type is a formula
		if (cType.equals(CellType.FORMULA))
		{
			// Set cell type to value type returned by formula
			cType = c.getCachedFormulaResultType();
		}
		
		// If cell is blank
		if (cType.equals(CellType.BLANK))
		{
			// Set return string to 'Empty Value'
			stringCValue = "Empty Value";
		}
		// If cell type is a String
		else if (cType.equals(CellType.STRING))
		{
			// Set return string to the string value
			stringCValue = c.getStringCellValue();
		}
		// If cell type is a number
		else if (cType.equals(CellType.NUMERIC))
		{
			// Set return string to String value of number
			stringCValue = NumberToTextConverter.toText(c.getNumericCellValue());
		}
		// If cell type is a boolean
		else if (cType.equals(CellType.BOOLEAN))
		{
			// Set return string to String value of boolean
			stringCValue = String.valueOf(c.getBooleanCellValue());
		}
		// If cell type is an error
		else if (cType.equals(CellType.ERROR))
		{
			// Set return string to 'Error Value'
			stringCValue = "Error Value";
		}
		// If cell type is none of the above
		else
		{
			// Set return string to 'Unknown Value'
			stringCValue = "Unknown Value";
		}
		
		// Return string
		return stringCValue;
	}
}
