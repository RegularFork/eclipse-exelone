package excelone;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.awt.event.ActionEvent;
import javax.swing.JProgressBar;
import javax.swing.JSeparator;
import javax.swing.JSpinner;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import java.awt.Color;
import java.awt.Window.Type;
import java.awt.Dialog.ModalExclusionType;
import javax.swing.UIManager;
import javax.swing.filechooser.FileFilter;

import com.formdev.flatlaf.FlatIntelliJLaf;
import com.raven.datechooser.DateChooser;
import com.raven.datechooser.listener.DateChooserAction;
import com.raven.datechooser.listener.DateChooserAdapter;

import java.awt.SystemColor;

public class Window {

	private JFrame frame;
	private JTextField textSourceField;
	private JTextField textTargetField;
	private JButton getTargetFileButton;
	private JButton cancelTargetFileButton;
	private JTextField textBreDataField;
	private JTextField textDataKegocField;
	private JTextField textKegocField;
	private static DateChooser chDate = new DateChooser();
	private static DateChooser chDate2 = new DateChooser();
	private final static String FILE_EXTENSION = ".xlsx"; 

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Window window = new Window();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Window() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		final ExcelService service = new ExcelService();
		
		
		
		frame = new JFrame("EXCELONE v2.0");
		frame.getContentPane().setFont(new Font("Arial Narrow", Font.PLAIN, 11));
		frame.setBounds(100, 100, 600, 500);
		frame.setResizable(false);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		textSourceField = new JTextField();
		textSourceField.setBackground(SystemColor.inactiveCaptionBorder);
		textSourceField.setText("C:\\Users\\commercial\\Documents\\РасходПоОбъектам1.xlsx");
		textSourceField.setEditable(false);
		textSourceField.setBounds(7, 86, 567, 20);
		frame.getContentPane().add(textSourceField);
		textSourceField.setColumns(10);
		
		JButton getSourceFileButton = new JButton("Выбрать источник данных");
		getSourceFileButton.setToolTipText("Выбрать файл \"РасходПоОбъектам\" для извлечения данных");
		getSourceFileButton.setBackground(UIManager.getColor("Button.light"));
		getSourceFileButton.setBounds(7, 107, 205, 23);
		getSourceFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setCurrentDirectory(new File("C:\\Users\\commercial\\Documents\\РасходПоОбъектам1.xlsx"));
				// fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

				fileChooser.setFileFilter(new FileFilter() {
					@Override
					public String getDescription() {
						return "Excel (.xlsx)";
					}

					@Override
					public boolean accept(File f) {
						if (f.getName().endsWith(FILE_EXTENSION) || f.isDirectory())
							return true;
						return false;
					}
				});

				fileChooser.showDialog(frame, "Выбрать файл");
				File f;
				try {
					f = fileChooser.getSelectedFile();
					textSourceField.setText(f.getAbsolutePath());
					service.fileToReadPath = f.getAbsolutePath();
					System.out.println(service.fileToReadPath);
				} catch (NullPointerException e1) {

					if (!textSourceField.getText().endsWith(FILE_EXTENSION)) {
						textSourceField.setText("Файл не выбран");
						System.out.println("Source File not choosed");
					}

				}
			}
		});
		frame.getContentPane().add(getSourceFileButton);
		
		JButton cancelSourceFileButton = new JButton("Х");
		cancelSourceFileButton.setToolTipText("Отменить выбранный файл");
		cancelSourceFileButton.setBackground(UIManager.getColor("CheckBox.focus"));
		cancelSourceFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		cancelSourceFileButton.setBounds(222, 107, 42, 23);
		frame.getContentPane().add(cancelSourceFileButton);
		
		textTargetField = new JTextField();
		textTargetField.setBackground(SystemColor.inactiveCaptionBorder);
		textTargetField.setText("\\\\172.16.16.16\\коммерческий отдел\\СЕНТЯБРЬ БРЭ ежедневный заполнять его!!!.xlsx");
		textTargetField.setEditable(false);
		textTargetField.setColumns(10);
		textTargetField.setBounds(7, 141, 567, 20);
		frame.getContentPane().add(textTargetField);
		
		getTargetFileButton = new JButton("Выбрать файл для записи");
		getTargetFileButton.setToolTipText("Выбрать файл \"Ежедневный БРЭ\" для записи данных");
		getTargetFileButton.setBackground(UIManager.getColor("Button.light"));
		getTargetFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setCurrentDirectory(new File("\\\\172.16.16.16\\коммерческий отдел\\СЕНТЯБРЬ БРЭ ежедневный заполнять его!!!.xlsx"));
				// fileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

				fileChooser.setFileFilter(new FileFilter() {
					@Override
					public String getDescription() {
						return "Excel (.xlsx)";
					}

					@Override
					public boolean accept(File f) {
						if (f.getName().endsWith(FILE_EXTENSION) || f.isDirectory())
							return true;
						return false;
					}
				});

				fileChooser.showDialog(frame, "Выбрать файл");
				File f;
				try {
					f = fileChooser.getSelectedFile();
					textTargetField.setText(f.getAbsolutePath());
					service.fileToWritePath = f.getAbsolutePath();
					System.out.println(service.fileToWritePath);
				} catch (NullPointerException e1) {

					if (!textTargetField.getText().endsWith(FILE_EXTENSION)) {
						textTargetField.setText("Файл не выбран");
						System.out.println("Target File not choosed");
					}

				}

			}
			
		});
		getTargetFileButton.setBounds(7, 162, 205, 23);
		frame.getContentPane().add(getTargetFileButton);
		
		cancelTargetFileButton = new JButton("Х");
		cancelTargetFileButton.setToolTipText("Отменить выбранный файл");
		cancelTargetFileButton.setBackground(UIManager.getColor("CheckBox.focus"));
		cancelTargetFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		cancelTargetFileButton.setBounds(222, 162, 42, 23);
		cancelTargetFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				textTargetField.setText("Пусто");
				service.fileToWritePath = "";
				System.out.println("Target reset");
			}
		});
		frame.getContentPane().add(cancelTargetFileButton);

		JProgressBar progressBar = new JProgressBar();
		progressBar.setBounds(10, 436, 564, 14);
		progressBar.setValue(30);
		frame.getContentPane().add(progressBar);
		
		JSeparator separator = new JSeparator();
		separator.setBounds(7, 257, 574, 2);
		frame.getContentPane().add(separator);
		
		final JComboBox hourFromCombo = new JComboBox();
		hourFromCombo.setBackground(SystemColor.inactiveCaptionBorder);
		hourFromCombo.setToolTipText("Указывается прошедший час");
		hourFromCombo.setModel(new DefaultComboBoxModel(new String[] {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24"}));
		hourFromCombo.setBounds(10, 53, 51, 22);
		System.out.println("firstHour " + service.firstHour);
		hourFromCombo.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
				int compareFirst = Integer.parseInt((String) hourFromCombo.getSelectedItem());
				if (compareFirst < service.currentHour) {
					service.firstHour = compareFirst;
					System.out.println("firstHour set to " + service.firstHour);
				} else {
				service.firstHour = service.currentHour;
				System.out.println("firstHour set to " + service.firstHour);
				hourFromCombo.setSelectedItem(String.valueOf(service.currentHour));
				}
			}
		});
		frame.getContentPane().add(hourFromCombo);
		
		final JComboBox hourUntilCombo = new JComboBox();
		hourUntilCombo.setBackground(SystemColor.inactiveCaptionBorder);
		hourUntilCombo.setToolTipText("Указывается прошедший час");
		hourUntilCombo.setModel(new DefaultComboBoxModel(new String[] {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24"}));
		hourUntilCombo.setBounds(81, 53, 51, 22);
		System.out.println("CurrentHour " + service.currentHour);
		hourUntilCombo.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
				int compareLast = Integer.parseInt((String) hourUntilCombo.getSelectedItem());
				if (compareLast > service.firstHour) {
				service.currentHour = compareLast;
				System.out.println("currentHour set to " + service.currentHour);
				} else {
					service.currentHour = service.firstHour;
					System.out.println("firstHour set to " + service.firstHour);
					hourUntilCombo.setSelectedItem(String.valueOf(service.firstHour));
				}
			}
		});
		frame.getContentPane().add(hourUntilCombo);
		
		JLabel hourFromLabel = new JLabel("От");
		hourFromLabel.setHorizontalAlignment(SwingConstants.CENTER);
		hourFromLabel.setBounds(10, 36, 46, 14);
		frame.getContentPane().add(hourFromLabel);
		
		JLabel hourUntilLabel = new JLabel("До");
		hourUntilLabel.setHorizontalAlignment(SwingConstants.CENTER);
		hourUntilLabel.setBounds(81, 36, 46, 14);
		frame.getContentPane().add(hourUntilLabel);
		
		textBreDataField = new JTextField();
		textBreDataField.setBackground(SystemColor.inactiveCaptionBorder);
		textBreDataField.setEditable(false);
		textBreDataField.setToolTipText("Дата снятия показаний");
		textBreDataField.setBounds(142, 53, 86, 22);
		chDate.setTextField(textBreDataField);
		service.currentDay = chDate.getSelectedDate().getDate();
		service.currentMonth = chDate.getSelectedDate().getMonth() + 1;
		service.currentYear = chDate.getSelectedDate().getYear() + 1900;
		System.out.println(service.currentDay + " / " + service.currentMonth + " / " + service.currentYear);
		chDate.setLabelCurrentDayVisible(false);
		chDate.addActionDateChooserListener(new DateChooserAdapter() {

			@Override
			public void dateChanged(Date date, DateChooserAction action) {
				// TODO Auto-generated method stub
				super.dateChanged(date, action);
				service.currentDay = date.getDate();
				service.currentMonth = date.getMonth() + 1;
				service.currentYear = date.getYear() + 1900;
				System.out.println(service.currentDay + " / " + service.currentMonth + " / " + service.currentYear);
			}

		});
		frame.getContentPane().add(textBreDataField);
		textBreDataField.setColumns(10);
		
		JLabel dataKEGOCLabel = new JLabel("Дата");
		dataKEGOCLabel.setBackground(UIManager.getColor("Button.light"));
		dataKEGOCLabel.setHorizontalAlignment(SwingConstants.CENTER);
		dataKEGOCLabel.setBounds(159, 36, 46, 14);
		frame.getContentPane().add(dataKEGOCLabel);
		
		JButton startCopyBreButton = new JButton("Копировать данные");
		startCopyBreButton.setBackground(UIManager.getColor("Button.light"));
		startCopyBreButton.setToolTipText("Запустить копирование данных");
		startCopyBreButton.setBounds(369, 162, 205, 23);
		frame.getContentPane().add(startCopyBreButton);
		
		JLabel titleBRE = new JLabel("Копирование данных в \"БРЭ-ежедневный\"");
		titleBRE.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		titleBRE.setBounds(10, 11, 264, 14);
		frame.getContentPane().add(titleBRE);
		
		JLabel titleBRE_1 = new JLabel("Создание файла \"БРЭ для КЕГОК\"");
		titleBRE_1.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		titleBRE_1.setBounds(7, 270, 264, 14);
		frame.getContentPane().add(titleBRE_1);
		
//		JLabel dataKEGOCLabel_1 = new JLabel("Дата");
//		dataKEGOCLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
//		dataKEGOCLabel_1.setBounds(28, 192, 46, 14);
//		frame.getContentPane().add(dataKEGOCLabel_1);
//		
//		textDataKegocField = new JTextField();
//		textDataKegocField.setBackground(SystemColor.inactiveCaptionBorder);
//		textDataKegocField.setToolTipText("Дата снятия показаний");
//		textDataKegocField.setEditable(false);
//		textDataKegocField.setColumns(10);
//		textDataKegocField.setBounds(11, 209, 86, 22);
//		chDate2.setTextField(textDataKegocField);
//		System.out.println(service.currentDay + " / " + service.currentMonth + " / " + service.currentYear);
//		chDate2.setLabelCurrentDayVisible(false);
//		chDate2.addActionDateChooserListener(new DateChooserAdapter() {
//
//			@Override
//			public void dateChanged(Date date, DateChooserAction action) {
//				// TODO Auto-generated method stub
//
//				service.currentDay = date.getDate();
//				service.currentMonth = date.getMonth() + 1;
//				service.currentYear = date.getYear() + 1900;
//				chDate.setTextField(textDataKegocField);
//				System.out.println(service.currentDay + " / " + service.currentMonth + " / " + service.currentYear);
//			}
//
//		});
//		frame.getContentPane().add(textDataKegocField);
		
		textKegocField = new JTextField();
		textKegocField.setBackground(SystemColor.inactiveCaptionBorder);
		textKegocField.setText("C:\\Users\\commercial\\Desktop\\Суточная ведомость\\" + "БРЭ для KEGOC Bassel " + service.currentDay + " сентября.xlsx");
		textKegocField.setEditable(false);
		textKegocField.setColumns(10);
		textKegocField.setBounds(7, 295, 566, 20);
		frame.getContentPane().add(textKegocField);
		
		JButton setKegocFolderButton = new JButton("Выбрать источник данных");
		setKegocFolderButton.setToolTipText("Выбрать файл \"РасходПоОбъектам\" для извлечения данных");
		setKegocFolderButton.setBackground(UIManager.getColor("Button.light"));
		setKegocFolderButton.setBounds(7, 319, 204, 23);
		setKegocFolderButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setCurrentDirectory(new File("C:\\Users\\commercial\\Desktop\\Суточная ведомость\\"));
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

//				fileChooser.setFileFilter(new FileFilter() {
//					@Override
//					public String getDescription() {
//						return "JAR";
//					}
//
//					@Override
//					public boolean accept(File f) {
//						if (f.getName().endsWith(FILE_EXTENSION) || f.isDirectory())
//							return true;
//						return false;
//					}
//				});

				fileChooser.showDialog(frame, "Выбрать папку");
				File f;
				try {
					f = fileChooser.getSelectedFile();
					textKegocField.setText(f.getAbsolutePath() + "\\БРЭ для KEGOC Bassel " + service.currentDay + " сентября.xlsx");
					service.fileToWriteDailyPath = f.getAbsolutePath();
					System.out.println(service.fileToWriteDailyPath);
				} catch (NullPointerException e1) {

					if (!textKegocField.getText().endsWith(FILE_EXTENSION)) {
						textKegocField.setText("Файл не выбран");
						System.out.println("Kegoc Folder not choosed");
					}

				}
			}
		});
		frame.getContentPane().add(setKegocFolderButton);
		
		JButton cancelKegocFileButton = new JButton("Х");
		cancelKegocFileButton.setToolTipText("Отменить выбранный файл");
		cancelKegocFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		cancelKegocFileButton.setBackground(UIManager.getColor("CheckBox.focus"));
		cancelKegocFileButton.setBounds(222, 319, 41, 23);
		frame.getContentPane().add(cancelKegocFileButton);
		
		JButton startCopyBreButton_1 = new JButton("Создать файл");
		startCopyBreButton_1.setToolTipText("Создание ежедневного файла \"БРЭ для КЕГОК\"");
		startCopyBreButton_1.setBackground(UIManager.getColor("Button.light"));
		startCopyBreButton_1.setBounds(370, 319, 204, 23);
		frame.getContentPane().add(startCopyBreButton_1);
		
		FlatIntelliJLaf.registerCustomDefaultsSource("style");
		FlatIntelliJLaf.setup();

		frame.revalidate();
		
	}
}
