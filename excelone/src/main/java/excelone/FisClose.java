package excelone;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class FisClose {

	public static void main(String[] args) throws IOException {
		System.out.println("FIS close");
		// TODO Auto-generated method stub
		try {
			FileInputStream fis = new FileInputStream("\\\\172.16.16.16\\коммерческий отдел\\СЕНТЯБРЬ БРЭ ежедневный заполнять его!!!.xlsx");
			fis.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
