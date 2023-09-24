package excelone;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {
	public static void main(String[] args) throws IOException {

		ExcelService service = new ExcelService();
//		service.copyOneRow();
//		service.currentDay = 21;
//		service.currentHour = 24;
//		service.copyAllRowsPerDay();
		service.setTimeToCopy();

		System.out.println("Готово. Вы великолепны!");
	}

}
