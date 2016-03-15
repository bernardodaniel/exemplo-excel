package teste;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Leitura {

	public static void main(String[] args) throws Exception {
		
		FileInputStream fis = new FileInputStream("teste.xlsx");
		
		Workbook workbook = new XSSFWorkbook(fis);
		
		Sheet sheet = workbook.getSheetAt(0);
		
		for (Row row : sheet) {
			Cell cell = row.getCell(0);
			if ("Fim".equalsIgnoreCase(cell.getStringCellValue())) {
				break;
			}
			
			cell = row.getCell(1);
			System.out.print(cell.getStringCellValue() + " | ");
			
			cell = row.getCell(2);
			System.out.print(cell.getStringCellValue() + " | ");
			
			cell = row.getCell(3);
			System.out.println(cell.getNumericCellValue());
		}
		
		workbook.close();

	}

}
