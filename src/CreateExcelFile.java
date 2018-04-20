import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class CreateExcelFile {
	

	public XSSFWorkbook writeXLSXFile() throws IOException {
		
		String excelFileName = "test.xlsx";

		String sheetName = "Sheet1"; 

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName) ;

		// Iterating r number of rows
		for (int r=0;r < 5; r++ ) {
			XSSFRow row = sheet.createRow(r);

			// Iterating c number of columns
			for (int c=0;c < 5; c++ )
			{
				XSSFCell cell = row.createCell(c);
	
				cell.setCellValue("Cell "+r+" "+c);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(excelFileName);

		// Write this workbook to an outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		
		return wb;
	}
	
}
