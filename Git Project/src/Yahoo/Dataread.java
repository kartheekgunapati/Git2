package Yahoo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Dataread {
	
	XSSFWorkbook wb;
	XSSFSheet sheet;
	
	public void excelData() throws Exception {
		try {
			File src = new File("C:\\Users\\msuser1\\git\\Git2\\Git Project\\TestData\\Yahoo Acct New.xlsx");
			FileInputStream fis = new FileInputStream(src);
			wb = new XSSFWorkbook(fis);
		} catch (FileNotFoundException e) {
			System.out.println(e.getMessage());
		}
		
	}
	
	public String userData() {
		sheet = wb.getSheetAt(0);
		String data = sheet.getRow(1).getCell(1).getStringCellValue();
		return data;
	}

	public int dataCount() {
		int row = wb.getSheetAt(0).getLastRowNum();
		row = row+1;
		return row;
	}
}
