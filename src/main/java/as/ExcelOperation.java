package as;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperation {

	
	//data writing
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create a blank sheet

		XSSFSheet sheet = workbook.createSheet("sheet1");

		//create row
		Row row = sheet.createRow(0);
		//Create cell
		Cell cell = row.createCell(0);
		//data
		cell.setCellValue("Satya");

		//sending file to local
		FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Admin\\OneDrive\\Desktop\\Demo.xlsx"));
		workbook.write(out);
		System.out.println("Data loaded");

	}
}
