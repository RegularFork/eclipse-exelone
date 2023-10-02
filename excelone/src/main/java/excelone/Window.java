package excelone;

import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.SystemColor;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JSeparator;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.filechooser.FileFilter;

import com.formdev.flatlaf.intellijthemes.*;

import com.raven.datechooser.DateChooser;
import com.raven.datechooser.listener.DateChooserAction;
import com.raven.datechooser.listener.DateChooserAdapter;
import javax.swing.JToggleButton;

public class Window {

	private JFrame frame;
	private JTextField textSourceField;
	private JTextField textTargetField;
	private JButton getTargetFileButton;
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
		FlatDraculaIJTheme.setup();
		final ExcelService service = new ExcelService();
		service.setCurrentDate();

		frame = new JFrame("ЛЕНИВАЯ ЖОПА v2.0");
		frame.getContentPane().setFont(new Font("Arial Narrow", Font.PLAIN, 11));
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Dimension dimension = toolkit.getScreenSize();
		frame.setBounds(dimension.width/2 - 300, dimension.height/2 - 260, 600, 520);
		frame.setResizable(false);
		frame.getContentPane().setLayout(null);

		textSourceField = new JTextField();
//		textSourceField.setBackground(SystemColor.inactiveCaptionBorder);
		textSourceField.setText(service.fileToReadPath);
		textSourceField.setEditable(false);
		textSourceField.setBounds(7, 86, 567, 20);
		frame.getContentPane().add(textSourceField);
		textSourceField.setColumns(10);

		final JLabel progressLabel = new JLabel("");
		progressLabel.setBounds(7, 199, 567, 22);
		frame.getContentPane().add(progressLabel);


//		FlatArcOrangeIJTheme.setup();

		
		
		//		FlatIntelliJLaf.registerCustomDefaultsSource("style");
//		FlatIntelliJLaf.setup();

		frame.revalidate();

		JButton getSourceFileButton = new JButton("Выбрать источник данных");
		getSourceFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		getSourceFileButton.setToolTipText("Выбрать файл \"РасходПоОбъектам\" для извлечения данных");
//		getSourceFileButton.setBackground(SystemColor.text);
		getSourceFileButton.setBounds(7, 107, 205, 23);
		getSourceFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setCurrentDirectory(new File(service.fileToReadPath));
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

		textTargetField = new JTextField();
//		textTargetField.setBackground(SystemColor.inactiveCaptionBorder);
		textTargetField.setText(service.fileToWritePath);
		textTargetField.setEditable(false);
		textTargetField.setColumns(10);
		textTargetField.setBounds(7, 141, 567, 20);
		frame.getContentPane().add(textTargetField);

		getTargetFileButton = new JButton("Выбрать файл для записи");
		getTargetFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		getTargetFileButton.setToolTipText("Выбрать файл \"Ежедневный БРЭ\" для записи данных");
//		getTargetFileButton.setBackground(SystemColor.text);
		getTargetFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setCurrentDirectory(new File(service.fileToWritePath));
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

//		final JProgressBar progressBar = new JProgressBar();
//		progressBar.setEnabled(false);
//		progressBar.setStringPainted(true);
//		progressBar.setIndeterminate(false);
//		progressBar.setBounds(7, 436, 567, 14);
//		progressBar.setValue(0);
//		frame.getContentPane().add(progressBar);

		JSeparator separator = new JSeparator();
		separator.setBounds(7, 232, 567, 2);
		frame.getContentPane().add(separator);

		final JComboBox hourFromCombo = new JComboBox();
		final JComboBox hourUntilCombo = new JComboBox();
//		hourFromCombo.setBackground(SystemColor.inactiveCaptionBorder);
		hourFromCombo.setToolTipText("Указывается прошедший час");
		hourFromCombo.setModel(new DefaultComboBoxModel(new String[] { "1", "2", "3", "4", "5", "6", "7", "8", "9",
				"10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24" }));
		hourFromCombo.setBounds(10, 53, 51, 22);
		System.out.println("firstHour " + service.firstHour);
		hourFromCombo.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				int compareFirst = Integer.parseInt((String) hourFromCombo.getSelectedItem());
				if (compareFirst < service.currentHour) {
					service.firstHour = compareFirst;
					System.out.println("firstHour set to " + service.firstHour);
				} else {
					service.firstHour = compareFirst;
					service.currentHour = compareFirst;
					System.out.println("firstHour set to " + service.firstHour);
					System.out.println("currenttHour up to " + service.currentHour);
					hourFromCombo.setSelectedItem(String.valueOf(service.firstHour));
					hourUntilCombo.setSelectedItem(String.valueOf(service.currentHour));

				}
			}
		});
		frame.getContentPane().add(hourFromCombo);

//		final JComboBox hourUntilCombo = new JComboBox();
//		hourUntilCombo.setBackground(SystemColor.inactiveCaptionBorder);
		hourUntilCombo.setToolTipText("Указывается прошедший час");
		hourUntilCombo.setModel(new DefaultComboBoxModel(new String[] { "1", "2", "3", "4", "5", "6", "7", "8", "9",
				"10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24" }));
		hourUntilCombo.setBounds(81, 53, 51, 22);
		System.out.println("CurrentHour " + service.currentHour);
		hourUntilCombo.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				int compareLast = Integer.parseInt((String) hourUntilCombo.getSelectedItem());
				if (compareLast > service.firstHour) {
					service.currentHour = compareLast;
					System.out.println("currentHour set to " + service.currentHour);
				} else {
					service.currentHour = compareLast;
					service.firstHour = compareLast;
					System.out.println("currentHour set to " + service.currentHour);
					System.out.println("firstHour down to " + service.firstHour);
					hourUntilCombo.setSelectedItem(String.valueOf(service.currentHour));
					hourFromCombo.setSelectedItem(String.valueOf(service.firstHour));
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
//		textBreDataField.setBackground(SystemColor.inactiveCaptionBorder);
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
				super.dateChanged(date, action);
				service.currentDay = date.getDate();
				service.currentMonth = date.getMonth() + 1;
				service.currentYear = date.getYear() + 1900;
				textKegocField.setText(
						service.fileToWriteDailyPath + "БРЭ для KEGOC Bassel " + service.currentDay + service.getMonthStringName() + ".xlsx");
				System.out.println(service.currentDay + " / " + service.currentMonth + " / " + service.currentYear);
			}

		});
		frame.getContentPane().add(textBreDataField);
		textBreDataField.setColumns(10);

		JLabel dataKEGOCLabel = new JLabel("Дата");
//		dataKEGOCLabel.setBackground(UIManager.getColor("Button.light"));
		dataKEGOCLabel.setHorizontalAlignment(SwingConstants.CENTER);
		dataKEGOCLabel.setBounds(159, 36, 46, 14);
		frame.getContentPane().add(dataKEGOCLabel);

		JButton startCopyBreButton = new JButton("Копировать данные");
		startCopyBreButton.setFont(new Font("Tahoma", Font.BOLD, 11));
//		startCopyBreButton.setBackground(SystemColor.inactiveCaption);
		startCopyBreButton.setToolTipText("Запустить копирование данных");
		startCopyBreButton.setBounds(369, 162, 205, 23);
		startCopyBreButton.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {

				try {
					progressLabel.setText("Копирование данных...");
					if (!service.compareDateAskue())
						JOptionPane.showMessageDialog(frame,
								("Различные даты в файлах\n\"" + new File(service.fileToReadPath).getName() + "\"\nи\n\""
										+ new File(service.fileToWritePath).getName() + "\""),
								"Ошибка", JOptionPane.ERROR_MESSAGE);
					else {
						int input = JOptionPane.showConfirmDialog(frame,
								("Скопировать данные в\n" + new File(service.fileToWritePath).getName()),
								"Подтвердить действие", JOptionPane.YES_NO_OPTION);
						if (input == 0) {
							frame.revalidate();
							try {
								service.copyAllRowsPerDay();
								progressLabel.setText("Выполнено!");
								JOptionPane.showMessageDialog(frame,
										("Скопированы данные за " + (service.currentHour - (service.firstHour - 1))
												+ service.getCorrechHourString()),
										"Копирование данных завершено", JOptionPane.INFORMATION_MESSAGE);
								System.out.println("Copy Complete!");
								System.out.println(input);
							} catch (FileNotFoundException e1) {
								// TODO Auto-generated catch block
								JOptionPane
										.showMessageDialog(frame,
												"Файл \"" + new File(service.fileToWritePath).getName()
														+ "\"\n недоступен для записи",
												"Ошибка", JOptionPane.ERROR_MESSAGE);
								System.out.println("Файл недоступен или открыт");
							}
						}
					}
					progressLabel.setText("");
				} catch (IllegalArgumentException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		frame.getContentPane().add(startCopyBreButton);

		JLabel titleBRE = new JLabel("Копирование данных в \"БРЭ-ежедневный\"");
		titleBRE.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		titleBRE.setBounds(10, 11, 264, 14);
		frame.getContentPane().add(titleBRE);

		JLabel titleBRE_1 = new JLabel("Создание файла \"БРЭ для КЕГОК\"");
		titleBRE_1.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		titleBRE_1.setBounds(7, 245, 264, 14);
		frame.getContentPane().add(titleBRE_1);

		textKegocField = new JTextField();
//		textKegocField.setBackground(SystemColor.inactiveCaptionBorder);
		textKegocField.setText(
				service.fileToWriteDailyPath + "БРЭ для KEGOC Bassel " + service.currentDay + service.getMonthStringName() + ".xlsx");
		textKegocField.setEditable(false);
		textKegocField.setColumns(10);
		textKegocField.setBounds(7, 268, 566, 20);
		frame.getContentPane().add(textKegocField);

		JButton setKegocFolderButton = new JButton("Выбрать папку");
		setKegocFolderButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		setKegocFolderButton.setToolTipText("Выбрать файл \"РасходПоОбъектам\" для извлечения данных");
//		setKegocFolderButton.setBackground(SystemColor.text);
		setKegocFolderButton.setBounds(7, 294, 204, 23);
		setKegocFolderButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setCurrentDirectory(new File(service.fileToWriteDailyPath));
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				fileChooser.showDialog(frame, "Выбрать папку");
				File f;
				try {
					f = fileChooser.getSelectedFile();
					textKegocField.setText(
							f.getAbsolutePath() + "\\БРЭ для KEGOC Bassel " + service.currentDay + service.getMonthStringName() + ".xlsx");
					service.fileToWriteDailyPath = f.getAbsolutePath() + "\\";
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

		JButton createFileBre = new JButton("Создать файл");
		createFileBre.setFont(new Font("Tahoma", Font.BOLD, 11));
		createFileBre.setToolTipText("Создание ежедневного файла \"БРЭ для КЕГОК\"");
//		createFileBre.setBackground(SystemColor.inactiveCaption);
		createFileBre.setBounds(370, 294, 204, 23);
		createFileBre.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				int input = JOptionPane.showConfirmDialog(frame,
						("Создать файл \n" + "\"БРЭ для KEGOC Bassel " + service.currentDay + service.getMonthStringName() + ".xlsx\"?"),
						"Подтверждение операции", JOptionPane.YES_NO_OPTION);
				if (input == 0) {
					try {
						service.copyDailyRows();
						JOptionPane.showMessageDialog(frame,
								("Файл " + "\"БРЭ для KEGOC Bassel " + service.currentDay + service.getMonthStringName() + ".xlsx\"\nуспешно создан!"),
								"Файл создан", JOptionPane.INFORMATION_MESSAGE);
					} catch (FileNotFoundException e1) {
						JOptionPane
								.showMessageDialog(frame,
										("Файл " + "\"БРЭ для KEGOC Bassel " + service.currentDay + service.getMonthStringName() + ".xlsx\"\nнедоступен"),
										"Ошибка", JOptionPane.ERROR_MESSAGE);
						e1.printStackTrace();
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			}
		});

		frame.getContentPane().add(createFileBre);
		
		JLabel maxLabel = new JLabel("Максимум");
		maxLabel.setHorizontalAlignment(SwingConstants.CENTER);
		maxLabel.setBounds(25, 390, 86, 14);
		frame.getContentPane().add(maxLabel);
		
		JLabel minLabel = new JLabel("Минимум");
		minLabel.setHorizontalAlignment(SwingConstants.CENTER);
		minLabel.setBounds(181, 390, 83, 14);
		frame.getContentPane().add(minLabel);
		
		JLabel avgLabel = new JLabel("Среднее");
		avgLabel.setHorizontalAlignment(SwingConstants.CENTER);
		avgLabel.setBounds(337, 390, 86, 14);
		frame.getContentPane().add(avgLabel);
		
		JLabel totalLabel = new JLabel("Всего");
		totalLabel.setHorizontalAlignment(SwingConstants.CENTER);
		totalLabel.setBounds(489, 390, 71, 14);
		frame.getContentPane().add(totalLabel);
		
		JSeparator separator_1 = new JSeparator();
		separator_1.setBounds(7, 327, 567, 2);
		frame.getContentPane().add(separator_1);
		
		JLabel statsDailyLabel = new JLabel("Статистика с начала суток до текущего часа");
		statsDailyLabel.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		statsDailyLabel.setBounds(10, 365, 337, 14);
		frame.getContentPane().add(statsDailyLabel);
		
		final JLabel maxValueLabel = new JLabel("0");
		maxValueLabel.setHorizontalAlignment(SwingConstants.CENTER);
		maxValueLabel.setFont(new Font("Tahoma", Font.PLAIN, 14));
		maxValueLabel.setBounds(25, 419, 86, 14);
		frame.getContentPane().add(maxValueLabel);
		
		final JLabel minValueLabel = new JLabel("0");
		minValueLabel.setHorizontalAlignment(SwingConstants.CENTER);
		minValueLabel.setFont(new Font("Tahoma", Font.PLAIN, 14));
		minValueLabel.setBounds(181, 415, 83, 14);
		frame.getContentPane().add(minValueLabel);
		
		final JLabel avgValueLabel = new JLabel("0");
		avgValueLabel.setHorizontalAlignment(SwingConstants.CENTER);
		avgValueLabel.setFont(new Font("Tahoma", Font.PLAIN, 14));
		avgValueLabel.setBounds(337, 415, 86, 14);
		frame.getContentPane().add(avgValueLabel);
		
		final JLabel totalValueLabel = new JLabel("0");
		totalValueLabel.setHorizontalAlignment(SwingConstants.CENTER);
		totalValueLabel.setFont(new Font("Tahoma", Font.PLAIN, 14));
		totalValueLabel.setBounds(489, 415, 71, 14);
		frame.getContentPane().add(totalValueLabel);
		
		JButton statsButton = new JButton("Получить статистику");
		statsButton.setFont(new Font("Tahoma", Font.BOLD, 11));
//		statsButton.setBackground(SystemColor.text);
		statsButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					double[] stats = service.getStatistic();
					maxValueLabel.setText(String.valueOf(Math.round(stats[0])));
					minValueLabel.setText(String.valueOf(Math.round(stats[1])));
					avgValueLabel.setText(String.valueOf(Math.round(stats[2])));
					totalValueLabel.setText(String.valueOf(Math.round(stats[3])));
				} catch (FileNotFoundException e1) {
					JOptionPane.showMessageDialog(frame, "Файл недоступен", "ОШИБКА", JOptionPane.ERROR_MESSAGE);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		statsButton.setBounds(7, 447, 205, 23);
		frame.getContentPane().add(statsButton);

	}
}
