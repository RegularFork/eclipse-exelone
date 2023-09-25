package excelone;
import java.io.IOException;


public class Main2 {
	public static void main(String[] args){

		ExcelService service = new ExcelService();

		try {
			service.chooseMode();
		} catch (IOException e) {
			System.out.println("Всё пошло не так!");;
		}

		if (service.isSuccess) {
			service.printEndMessageSuccess();
		} else {
			service.printEndMessageFail();
		}
		
	}

}
