package excelone;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.GregorianCalendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelService {
	private final int SKIP_ROWS_SBRE = 4;
	private final int SKIP_ROWS_ASKUE = 10;
	private final int HOURS_IN_DAY = 24;
	public int currentDay;
	public int currentHour;
	private int currentRow = 32;
	private int rowTarget = 507;
	private int currentMonth = 9;
	String sourceFile;
	String targetFile;
	private String fileToReadPath = "C:\\Users\\commercial\\Documents\\РасходПоОбъектам1.xlsx";
	// Test file
//	private String fileToWritePath = "C:\\Users\\commercial\\Documents\\MyFiles\\excelpoint\\СЕНТЯБРЬ БРЭ result.xlsx";
	// Original file
	private String fileToWritePath = "\\\\172.16.16.16\\коммерческий отдел\\СЕНТЯБРЬ БРЭ ежедневный заполнять его!!!.xlsx";
	private String fileToWriteDailyTemplate = "C:\\Users\\commercial\\Documents\\MyFiles\\excelpoint\\БРЭ для KEGOC Bassel template.xlsx";
	private String fileToWriteDailyPath = "C:\\Users\\commercial\\Desktop\\Суточная ведомость\\";
	int[] cellNumbersToReadSBRE = { 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34,
			35, 37, 39, 41, 47, 49, 42, 44 };
	
	private void copyAllRowsPerDay() {
		for (int i = 0; i < currentHour; i++) {
			currentRow = SKIP_ROWS_ASKUE + i;
			rowTarget = ((currentDay - 1) * HOURS_IN_DAY + SKIP_ROWS_SBRE) + (i);
			copyOneRow();
			System.out.println("День " + currentDay + "; Час " + i + " - " + (i + 1));
			
		}
		System.out.println("Done!");
	}
	
	
	private void copyOneRow() {
		try {
			double[] rowValues = readRowFromPrimaryFile();
			writeRowToFile(rowValues);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// Чтение массива из АСКУЭ-файла
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
	// Запись массива в Ежедневный БРЭ
	private void writeRowToFile(double[] rowValues) throws IOException {
		Cell currentCell = null;
		System.out.println("Запись в строку " + rowTarget);
		FileInputStream fis = new FileInputStream(fileToWritePath);
		Workbook wbToWrite = WorkbookFactory.create(fis);		
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
        FileOutputStream fos = new FileOutputStream(new File(fileToWritePath));
        wbToWrite.write(fos);
        fos.close();
        wbToWrite.close();
	}
	

	// ======== Создание БРЭ для КЕГОК =====================================================================
	
	private void copyDailyRows() throws IOException {
		setCurrentDay();
		sourceFile = fileToWritePath;
		targetFile = fileToWriteDailyPath  + "БРЭ для KEGOC " + currentDay + " сентября.xlsx";
		try {
		double[][] dailyRows = readDailyRows();
		writeDailyValues(dailyRows);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	// Чтение данных за прошедшие сутки
	private double [][] readDailyRows() throws IOException{
		
		FileInputStream fis = new FileInputStream(sourceFile);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet = wb.getSheetAt(0);
		double[][]dailyRows = new double[24][30];
		for (int rows = 0; rows < 24; rows++) {
			int targetRow = ((currentDay - 1) * HOURS_IN_DAY + SKIP_ROWS_SBRE) + (rows);
			for (int columns = 0; columns < 28; columns++) {
				dailyRows[rows][columns] = sheet.getRow(targetRow).getCell(columns + 11).getNumericCellValue();
			}
			dailyRows[rows][28] = sheet.getRow(targetRow).getCell(5).getNumericCellValue();
			dailyRows[rows][29] = sheet.getRow(targetRow).getCell(6).getNumericCellValue();
		}
		fis.close();
		wb.close();
		System.out.println("Данные считаны");
		return dailyRows;
	}
	
	// Запись данных в новый файл БРЭ для КЕГОК
	private void writeDailyValues(double[][]dailyRows) throws IOException {
		FileInputStream fis = new FileInputStream(fileToWriteDailyTemplate);
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sheet = wb.getSheetAt(0);
		sheet.setForceFormulaRecalculation(true);
		for (int rows = 0; rows < 24; rows++) {
			for (int columns = 0; columns < 28; columns++) {
				sheet.getRow(rows + 4).getCell(columns + 5).setCellValue(dailyRows[rows][columns]);
			}
			sheet.getRow(rows + 4).getCell(2).setCellValue(dailyRows[rows][28]);
			sheet.getRow(rows + 4).getCell(3).setCellValue(dailyRows[rows][29]);
			System.out.println("Записаны данные часа " + rows + " - " + (rows + 1));
		}
		sheet.getRow(0).getCell(0).setCellValue(new GregorianCalendar(2023, currentMonth - 1, currentDay));
		fis.close();
		
		FileOutputStream fos = new FileOutputStream(targetFile);
		wb.write(fos);
		wb.close();
		fos.close();
	}
	
	//=========== Выбор режима порграммы =============================================================================
	
	public void chooseMode() throws IOException {
		BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
		System.out.println("Ленивая Жопа V1.1" + "\n==================\n");
			System.out.println("Выбрать действие:\n1 - Копирование данных\n2 - Создание файла для КЕГОК\n");
			int mode = Integer.parseInt(reader.readLine());
			System.out.println("Выбран режим  " + mode);
			if (mode == 1) {
				setTimeToCopy();
			}
			else if (mode == 2) {
				copyDailyRows();				
			}
			else {
				System.out.println("Ошибка ввода. Программа завершена");
		}
	}
	
	// Сетап для копирования показаний
	private void setTimeToCopy() throws IOException {
		BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
		System.out.println("====================\n"
				+ "**  ВНИМАНИЕ: должен быть загружен и сохранён актуальный файл \"РасходПоОбъектам1.xlsx\""
				+ "\n**  ОСОБОЕ ВНИМАНИЕ: неверно введённая дата перезапишет существующие данные\n\n"
				+ "За какое число нужно снять показания? (0 - 31):\n");
		currentDay = Integer.parseInt(reader.readLine());
		if (currentDay < 1 && currentDay > 31) {
			System.out.println("Неверно выбрана дата: " + currentDay);
			return;
		}
		System.out.println("За какой час снять показания? (0 - с начала суток)\n");
		currentHour = Integer.parseInt(reader.readLine());
		if (currentHour < 0 && currentHour > 24) {
			System.out.println("Неверно выбрано время: " + currentHour);
			return;
		}
		if (currentHour == 0) {
			System.out.println("До какого часа скопировать показания? ( 0 - до конца суток)\n");
			int input = Integer.parseInt(reader.readLine());
			if (input > 0 && input < 25) {
				currentHour = input;
				copyAllRowsPerDay();
			} else {
				System.out.println("Неверно выбрано время: " + currentHour);
				return;
			}
		} else {
			currentRow = SKIP_ROWS_ASKUE + currentHour - 1;
			rowTarget = ((currentDay - 1) * HOURS_IN_DAY + SKIP_ROWS_SBRE) + (currentHour - 1);
			copyOneRow();
		}
	}
	
	// Сетап для создания БРЭ для КЕГОК
	private void setCurrentDay() throws IOException {
		BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
		System.out.println("Введите число месяца для создания файла \"БРЭ для КЕГОК\"");
		int input = Integer.parseInt(reader.readLine());
		if (input < 1 && input > 31) {
			System.out.println("Неверно введена дата\nПрограмма завершена");
			return;
		} else currentDay = input;
		
	}
	
	

}
