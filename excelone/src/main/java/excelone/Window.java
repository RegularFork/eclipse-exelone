package excelone;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.Font;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.JProgressBar;
import javax.swing.JSeparator;
import javax.swing.JSpinner;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import java.awt.Color;
import java.awt.Window.Type;
import java.awt.Dialog.ModalExclusionType;
import javax.swing.UIManager;
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
		frame.setType(Type.POPUP);
		frame.getContentPane().setFont(new Font("Arial Narrow", Font.PLAIN, 11));
		frame.setBounds(100, 100, 600, 393);
		frame.setResizable(false);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		textSourceField = new JTextField();
		textSourceField.setText("Файл не выбран");
		textSourceField.setEditable(false);
		textSourceField.setBounds(7, 30, 308, 20);
		frame.getContentPane().add(textSourceField);
		textSourceField.setColumns(10);
		
		JButton getSourceFileButton = new JButton("Выбрать источник данных");
		getSourceFileButton.setToolTipText("Выбрать файл \"РасходПоОбъектам\" для извлечения данных");
		getSourceFileButton.setBackground(UIManager.getColor("CheckBox.background"));
		getSourceFileButton.setBounds(321, 29, 205, 23);
		frame.getContentPane().add(getSourceFileButton);
		
		JButton cancelSourceFileButton = new JButton("Х");
		cancelSourceFileButton.setToolTipText("Отменить выбранный файл");
		cancelSourceFileButton.setBackground(UIManager.getColor("CheckBox.focus"));
		cancelSourceFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		cancelSourceFileButton.setBounds(532, 29, 42, 23);
		frame.getContentPane().add(cancelSourceFileButton);
		
		textTargetField = new JTextField();
		textTargetField.setText("Файл не выбран");
		textTargetField.setEditable(false);
		textTargetField.setColumns(10);
		textTargetField.setBounds(7, 62, 308, 20);
		frame.getContentPane().add(textTargetField);
		
		getTargetFileButton = new JButton("Выбрать файл для записи");
		getTargetFileButton.setToolTipText("Выбрать файл \"Ежедневный БРЭ\" для записи данных");
		getTargetFileButton.setBackground(UIManager.getColor("CheckBox.background"));
		getTargetFileButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		getTargetFileButton.setBounds(321, 61, 205, 23);
		frame.getContentPane().add(getTargetFileButton);
		
		cancelTargetFileButton = new JButton("Х");
		cancelTargetFileButton.setToolTipText("Отменить выбранный файл");
		cancelTargetFileButton.setBackground(UIManager.getColor("CheckBox.focus"));
		cancelTargetFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		cancelTargetFileButton.setBounds(532, 61, 42, 23);
		frame.getContentPane().add(cancelTargetFileButton);
		
		JProgressBar progressBar = new JProgressBar();
		progressBar.setBounds(10, 236, 564, 14);
		frame.getContentPane().add(progressBar);
		
		JSeparator separator = new JSeparator();
		separator.setBounds(7, 135, 574, 2);
		frame.getContentPane().add(separator);
		
		final JComboBox hourFromCombo = new JComboBox();
		hourFromCombo.setBackground(SystemColor.inactiveCaptionBorder);
		hourFromCombo.setToolTipText("Указывается прошедший час");
		hourFromCombo.setModel(new DefaultComboBoxModel(new String[] {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24"}));
		hourFromCombo.setBounds(10, 102, 51, 22);
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
		hourUntilCombo.setBounds(81, 102, 51, 22);
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
		hourFromLabel.setBounds(10, 85, 46, 14);
		frame.getContentPane().add(hourFromLabel);
		
		JLabel hourUntilLabel = new JLabel("До");
		hourUntilLabel.setHorizontalAlignment(SwingConstants.CENTER);
		hourUntilLabel.setBounds(81, 85, 46, 14);
		frame.getContentPane().add(hourUntilLabel);
		
		textBreDataField = new JTextField();
		textBreDataField.setBackground(SystemColor.inactiveCaptionBorder);
		textBreDataField.setEditable(false);
		textBreDataField.setToolTipText("Дата снятия показаний");
		textBreDataField.setBounds(142, 102, 86, 22);
		frame.getContentPane().add(textBreDataField);
		textBreDataField.setColumns(10);
		
		JLabel dataKEGOCLabel = new JLabel("Дата");
		dataKEGOCLabel.setBackground(UIManager.getColor("Button.light"));
		dataKEGOCLabel.setHorizontalAlignment(SwingConstants.CENTER);
		dataKEGOCLabel.setBounds(159, 85, 46, 14);
		frame.getContentPane().add(dataKEGOCLabel);
		
		JButton startCopyBreButton = new JButton("Копировать данные");
		startCopyBreButton.setBackground(UIManager.getColor("Button.light"));
		startCopyBreButton.setToolTipText("Запустить копирование данных");
		startCopyBreButton.setBounds(321, 102, 205, 23);
		frame.getContentPane().add(startCopyBreButton);
		
		JLabel titleBRE = new JLabel("Копирование данных в \"БРЭ-ежедневный\"");
		titleBRE.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		titleBRE.setBounds(10, 11, 264, 14);
		frame.getContentPane().add(titleBRE);
		
		JLabel titleBRE_1 = new JLabel("Создание файла \"БРЭ для КЕГОК\"");
		titleBRE_1.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 13));
		titleBRE_1.setBounds(7, 148, 264, 14);
		frame.getContentPane().add(titleBRE_1);
		
		JLabel dataKEGOCLabel_1 = new JLabel("Дата");
		dataKEGOCLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
		dataKEGOCLabel_1.setBounds(28, 192, 46, 14);
		frame.getContentPane().add(dataKEGOCLabel_1);
		
		textDataKegocField = new JTextField();
		textDataKegocField.setBackground(SystemColor.inactiveCaptionBorder);
		textDataKegocField.setToolTipText("Дата снятия показаний");
		textDataKegocField.setEditable(false);
		textDataKegocField.setColumns(10);
		textDataKegocField.setBounds(11, 209, 86, 22);
		frame.getContentPane().add(textDataKegocField);
		
		textKegocField = new JTextField();
		textKegocField.setText("Файл не выбран");
		textKegocField.setEditable(false);
		textKegocField.setColumns(10);
		textKegocField.setBounds(8, 166, 307, 20);
		frame.getContentPane().add(textKegocField);
		
		JButton SetKegocFolderButton = new JButton("Выбрать источник данных");
		SetKegocFolderButton.setToolTipText("Выбрать файл \"РасходПоОбъектам\" для извлечения данных");
		SetKegocFolderButton.setBackground(UIManager.getColor("CheckBox.background"));
		SetKegocFolderButton.setBounds(322, 165, 204, 23);
		frame.getContentPane().add(SetKegocFolderButton);
		
		JButton cancelKegocFileButton = new JButton("Х");
		cancelKegocFileButton.setToolTipText("Отменить выбранный файл");
		cancelKegocFileButton.setFont(new Font("Tahoma", Font.BOLD, 11));
		cancelKegocFileButton.setBackground(UIManager.getColor("CheckBox.focus"));
		cancelKegocFileButton.setBounds(533, 165, 41, 23);
		frame.getContentPane().add(cancelKegocFileButton);
		
		JButton startCopyBreButton_1 = new JButton("Создать файл");
		startCopyBreButton_1.setToolTipText("Создание ежедневного файла \"БРЭ для КЕГОК\"");
		startCopyBreButton_1.setBackground(UIManager.getColor("CheckBox.light"));
		startCopyBreButton_1.setBounds(322, 209, 204, 23);
		frame.getContentPane().add(startCopyBreButton_1);
		
	}
}
