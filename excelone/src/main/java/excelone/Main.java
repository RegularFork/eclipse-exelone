package excelone;
import java.io.IOException;


public class Main {
	public static void main(String[] args){

		ExcelService service = new ExcelService();

		try {
			service.chooseMode();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		service.printEndMessageFail();
	}

}
