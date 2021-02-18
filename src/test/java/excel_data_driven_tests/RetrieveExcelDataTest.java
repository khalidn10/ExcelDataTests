package excel_data_driven_tests;

import java.io.IOException;
import java.util.ArrayList;

public class RetrieveExcelDataTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		RetrieveExcelData excelData = new RetrieveExcelData();
		ArrayList<String> data = excelData.getData("live data", "test case name", "purchase");
		
		// For each element in dataList array list
		for (String d : data)
		{
			// Print out data
			System.out.println(d);
		}
	}

}
