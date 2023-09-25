package excelone;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;


public class Main2 {
	public static void main(String[] args) throws IOException{
		
		System.out.println("====================================================\n"
				+ "= Добро пожаловать в программу \"Ленивая Жопа V1.3\" =\n"
				+ "====================================================\n");
		
            String[] command = {"C:\\Users\\commercial\\Desktop\\test.xlsx"};


		ExcelService service = new ExcelService();

		while(true) {
			try {
				service.chooseMode();
			} catch (IllegalArgumentException e) {
				System.out.println("Всё сломалось!");
			} catch (IOException e) {
				System.out.println("Всё пошло не так!");;
			}

			if (service.isSuccess) {
				service.printEndMessageSuccess();
			} else {
				service.printEndMessageFail();
			}
			
			System.out.println("\n\nЧто-нибудь ещё?\n\n\n");
		}
		
	}

}
