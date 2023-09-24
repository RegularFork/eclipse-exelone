package excelone;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelService {
	private final int SKIP_ROWS_SBRE = 4;
	private final int SKIP_ROWS_ASKUE = 10;
	private final int HOURS_IN_DAY = 24;
	public int currentDay;
	public int currentHour;
	private int currentRow = 32;
	private int rowTarget = 507;
	private String fileToReadPath = "C:\\Users\\commercial\\Documents\\MyFiles\\excelpoint\\РасходПоОбъектам1.xlsx";
	private String fileToWritePath = "C:\\Users\\commercial\\Documents\\MyFiles\\excelpoint\\СЕНТЯБРЬ БРЭ шаблон.xlsx";
	private String fileToWritePastDayPath = "C:\\Users\\commercial\\Desktop\\Суточная ведомость\\БРЭ для KEGOC Bassel 23 сентября.xlsx";
	private String fileToWritePastDayFinalPath = "C:\\Users\\commercial\\Desktop\\Суточная ведомость\\БРЭ для KEGOC Bassel 24 сентября.xlsx";
	private String fileToWriteFinalPath = "C:\\Users\\commercial\\Documents\\MyFiles\\excelpoint\\СЕНТЯБРЬ БРЭ result.xlsx";
	int[] cellNumbersToReadSBRE = { 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34,
			35, 37, 39, 41, 47, 49, 42, 44 };
	
	public void copyAllRowsPerDay() {
		for (int i = 0; i < currentHour; i++) {
			currentRow = SKIP_ROWS_ASKUE + i;
			rowTarget = ((currentDay - 1) * HOURS_IN_DAY + SKIP_ROWS_SBRE) + (i);
			copyOneRow();
			System.out.println("Write day " + currentDay + "; Hour " + i + " - " + (i + 1));
			
		}
		System.out.println("Done!");
	}
	
	
	public void copyOneRow() {
		try {
			double[] rowValues = readRowFromPrimaryFile();
			writeRowToFile(fileToWritePath, rowValues);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private double[] readRowFromPrimaryFile() throws IOException {

		double[] readedValue = new double[cellNumbersToReadSBRE.length];
		FileInputStream fileToRead = new FileInputStream(fileToReadPath);
		Workbook wbToRead = new XSSFWorkbook(fileToRead);
		Row rowToWrite = wbToRead.getSheetAt(0).getRow(currentRow);
		for (int i = 0; i < 30; i++) {
			readedValue[i] = rowToWrite.getCell(cellNumbersToReadSBRE[i]).getNumericCellValue();
		}
		fileToRead.close();
		wbToRead.close();
		return readedValue;
	}
	private void writeRowToFile(String fileToWritePath, double[] rowValues) throws IOException {
		Cell currentCell = null;
		System.out.println("Запись в строку " + rowTarget);
		FileInputStream fis = new FileInputStream(fileToWritePath);
		Workbook wbToWrite = new XSSFWorkbook(fileToWritePath);		
		Sheet sheetToWrite = wbToWrite.getSheetAt(0);
		sheetToWrite.setForceFormulaRecalculation(true);
        Row currentRow = sheetToWrite.getRow(rowTarget);
        for (int i = 0; i < 28; i++) {
        	currentCell = currentRow.getCell(i + 11);
        	currentCell.setCellValue(rowValues[i]);
        }
        currentRow.getCell(5).setCellValue(rowValues[28]);
        currentRow.getCell(6).setCellValue(rowValues[29]);
        fis.close();
        FileOutputStream fos = new FileOutputStream(new File(fileToWriteFinalPath));
        wbToWrite.write(fos);
        fos.close();
        wbToWrite.close();
	}
	
	public void setTimeToCopy() {
		Scanner scanner = new Scanner(System.in);
		System.out.println("За какое число нужно снять показания?");
		currentDay = scanner.nextInt();
		if (currentDay < 1 && currentDay > 31) {
			System.out.println("Неверно выбрана дата: " + currentDay);
			scanner.close();
			return;
		}
		System.out.println("За какой час снять показания? (0 - с начала суток)");
		currentHour = scanner.nextInt();
		if (currentHour < 0 && currentHour > 24) {
			System.out.println("Неверно выбрано время: " + currentHour);
			scanner.close();
			return;
		}
		if (currentHour == 0) {
			currentHour = 24;
			copyAllRowsPerDay();
			scanner.close();
		} else {
			currentRow = SKIP_ROWS_ASKUE + currentHour - 1;
			rowTarget = ((currentDay - 1) * HOURS_IN_DAY + SKIP_ROWS_SBRE) + (currentHour - 1);
			copyOneRow();
			scanner.close();
		}
	}
	

}