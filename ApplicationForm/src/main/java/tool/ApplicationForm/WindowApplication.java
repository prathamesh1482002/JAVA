package tool.ApplicationForm;

import java.awt.*;
import javax.swing.*;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.text.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;
import java.awt.event.ActionListener;
import java.io.*;
import java.awt.event.ActionEvent;

public class WindowApplication {

	private JFrame frame;
	private JTextField textField;
	private JTextField textField_1;
	private JTextField textField_2;
	private JTextField textField_3;
	private JTextField textField_4;
	private JTextField textField_5;
	private JTextField textField_6;
	private JTextField textField_7;
	private JTextField textField_8;
	private JTextPane textPane_JobDescription;
	private JTable table;
	private DefaultTableModel tableModel;
	private int totalRowCount;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					WindowApplication window = new WindowApplication();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public WindowApplication() {
		initialize();
	}

	private void initialize() {
		frame = new JFrame();
		frame.setResizable(false);
		frame.setFont(new Font("Dialog", Font.BOLD, 12));
		frame.setTitle("Application Form");
		frame.setBounds(100, 100, 658, 651);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		frame.setLocation(50, 30);

		JPanel panel = new JPanel();
		panel.setLayout(null);
		JScrollPane scrollPane = new JScrollPane(panel);
		scrollPane.setBounds(0, 0, 644, 612);
		scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		scrollPane.getVerticalScrollBar().setUnitIncrement(50);
		scrollPane.getVerticalScrollBar().setBlockIncrement(50);
		frame.getContentPane().add(scrollPane);
		panel.setPreferredSize(new Dimension(310, 920));

		textField = new JTextField();
		textField.setBounds(25, 158, 552, 20);
		panel.add(textField);
		textField.setColumns(10);

		textField_1 = new JTextField();
		textField_1.setBounds(25, 214, 552, 20);
		panel.add(textField_1);
		textField_1.setColumns(10);

		textField_2 = new JTextField();
		textField_2.setBounds(25, 270, 552, 20);
		panel.add(textField_2);
		textField_2.setColumns(10);

		JLabel lblNewLabel = new JLabel("Input 1");
		lblNewLabel.setBounds(25, 133, 61, 14);
		panel.add(lblNewLabel);

		JLabel lblNewLabel_1 = new JLabel("Input 2");
		lblNewLabel_1.setBounds(25, 189, 46, 14);
		panel.add(lblNewLabel_1);

		JLabel lblNewLabel_2 = new JLabel("Input 3");
		lblNewLabel_2.setBounds(25, 245, 61, 14);
		panel.add(lblNewLabel_2);

		JLabel lblNewLabel_3 = new JLabel("Input 4");
		lblNewLabel_3.setBounds(25, 301, 46, 14);
		panel.add(lblNewLabel_3);

		JLabel lblNewLabel_4 = new JLabel("Input 6");
		lblNewLabel_4.setBounds(25, 411, 86, 14);
		panel.add(lblNewLabel_4);

		textField_3 = new JTextField();
		textField_3.setBounds(25, 436, 552, 20);
		panel.add(textField_3);
		textField_3.setColumns(10);

		JLabel lblNewLabel_5 = new JLabel("Input 7");
		lblNewLabel_5.setBounds(25, 467, 109, 14);
		panel.add(lblNewLabel_5);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
		scrollPane_1.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane_1.setBounds(25, 492, 309, 100);
		panel.add(scrollPane_1);

		textPane_JobDescription = new JTextPane();
		scrollPane_1.setViewportView(textPane_JobDescription);

		JButton btnBold = new JButton("B");
		btnBold.setFont(new Font("Arial", Font.BOLD, 12));
		btnBold.setBounds(344, 492, 52, 25);
		panel.add(btnBold);

		JButton btnItalic = new JButton("I");
		btnItalic.setFont(new Font("Arial", Font.ITALIC, 12));
		btnItalic.setBounds(463, 492, 52, 25);
		panel.add(btnItalic);

		JButton btnUnderline = new JButton("U");
		btnUnderline.setFont(new Font("Arial", Font.PLAIN, 12));
		btnUnderline.setBounds(401, 492, 52, 25);
		panel.add(btnUnderline);

		JButton btnPlainText = new JButton("P");
		btnPlainText.setFont(new Font("Arial", Font.PLAIN, 12));
		btnPlainText.setBounds(525, 492, 52, 25);
		panel.add(btnPlainText);

		tableModel = new DefaultTableModel();

		btnBold.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("Bold", null);
					StyleConstants.setBold(style, true);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});

		btnItalic.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("Italic", null);
					StyleConstants.setItalic(style, true);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});
		btnUnderline.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("Underline", null);
					StyleConstants.setUnderline(style, true);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});
		JButton btnNewButton = new JButton("New Entry");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (!textField.getText().isEmpty() && !textField_1.getText().isEmpty()
						&& !textField_2.getText().isEmpty()
						&& !textField_3.getText().isEmpty() && !textField_4.getText().isEmpty()
						&& !textField_5.getText().isEmpty() && !textPane_JobDescription.getText().isEmpty()) {
					String filePath = textField_7.getText();
					String data1 = textField.getText();
					String data2 = textField_1.getText();
					String data3 = textField_2.getText();
					String data4 = textField_4.getText();
					String data5 = textField_5.getText();
					String data6 = textField_3.getText();
					String data7 = textPane_JobDescription.getText();
					StyledDocument doc = textPane_JobDescription.getStyledDocument();
					if (filePath.endsWith(".xlsx") || filePath.endsWith(".xls")) {
						writeToExcel(filePath, data1, data2, data3, data4, data5, data6, data7, doc);
						JOptionPane.showMessageDialog(frame, "Data Added Successfully");
						textField.setText("");
						textField_1.setText("");
						textField_2.setText("");
						textField_5.setText("");
						textField_3.setText("");
						textField_4.setText("");
						textPane_JobDescription.setText("");

						previewExcelFile(filePath);

					} else {
						JOptionPane.showMessageDialog(frame, "Please Select the Excel File");
					}
				} else {
					JOptionPane.showMessageDialog(frame, "Please Filled the Remainig Entries");
				}
			}
		});
		btnNewButton.setBounds(495, 877, 125, 23);
		panel.add(btnNewButton);

		JLabel lblNewLabel_7 = new JLabel("Input 5");
		lblNewLabel_7.setBounds(25, 357, 45, 13);
		panel.add(lblNewLabel_7);

		textField_5 = new JTextField();
		textField_5.setBounds(25, 381, 552, 19);
		panel.add(textField_5);
		textField_5.setColumns(10);

		JLabel lblNewLabel_9 = new JLabel("Select the Excel File");
		lblNewLabel_9.setBounds(25, 23, 125, 13);
		panel.add(lblNewLabel_9);

		textField_7 = new JTextField();
		textField_7.setBounds(25, 47, 447, 19);
		panel.add(textField_7);
		textField_7.setColumns(10);

		JButton btnNewButton_1 = new JButton("Browse");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select the Excel file");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField_7.setText(chooser.getSelectedFile().toString());
					String filePath = textField_7.getText();

					previewExcelFile(filePath);
				} else {
					JOptionPane.showMessageDialog(textField_7, "Invalid Path");
				}
			}
		});
		btnNewButton_1.setBounds(503, 47, 78, 21);
		panel.add(btnNewButton_1);
///////////FontButtons
		JButton btnNewButton_2 = new JButton("H1");
		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("H1", null);
					StyleConstants.setFontSize(style, 24);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});
		btnNewButton_2.setBounds(344, 569, 52, 23);
		panel.add(btnNewButton_2);

		JButton btnNewButton_2_1 = new JButton("H2");
		btnNewButton_2_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("H2", null);
					StyleConstants.setFontSize(style, 20);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});
		btnNewButton_2_1.setBounds(401, 569, 54, 23);
		panel.add(btnNewButton_2_1);

		JButton btnNewButton_2_1_1 = new JButton("H3");
		btnNewButton_2_1_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("H3", null);
					StyleConstants.setFontSize(style, 18);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});
		btnNewButton_2_1_1.setBounds(463, 569, 52, 23);
		panel.add(btnNewButton_2_1_1);

		JButton btnNewButton_2_1_1_1 = new JButton("H4");
		btnNewButton_2_1_1_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("H4", null);
					StyleConstants.setFontSize(style, 16);
					doc.setCharacterAttributes(start, end - start, style, false);

				}
			}
		});
		btnPlainText.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				StyledDocument doc = textPane_JobDescription.getStyledDocument();
				int start = textPane_JobDescription.getSelectionStart();
				int end = textPane_JobDescription.getSelectionEnd();
				if (start != end) {
					Style style = textPane_JobDescription.addStyle("Plain", null);
					StyleConstants.setBold(style, false);
					StyleConstants.setItalic(style, false);
					StyleConstants.setUnderline(style, false);
					StyleConstants.setFontSize(style, 12);
					doc.setCharacterAttributes(start, end - start, style, false);
				}
			}
		});
		btnNewButton_2_1_1_1.setBounds(525, 569, 52, 23);
		panel.add(btnNewButton_2_1_1_1);

		JLabel lblNewLabel_8 = new JLabel("Current RowNumber/Total Row Number");
		lblNewLabel_8.setBounds(25, 612, 227, 13);
		panel.add(lblNewLabel_8);

		textField_6 = new JTextField();
		textField_6.setBounds(344, 608, 72, 19);
		panel.add(textField_6);
		textField_6.setColumns(10);
		textField_8 = new JTextField();
		textField_8.setBounds(25, 102, 447, 20);
		panel.add(textField_8);
		textField_8.setColumns(10);

		JButton btnNewButton_3 = new JButton("Browse");
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select the PDF file");
				chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					textField_8.setText(chooser.getSelectedFile().toString());
					// String Pdfpath = textField_8.getText();
					File file = new File(textField_8.getText());
					if (textField_8.getText().endsWith(".pdf")) {
						Pdf.openFile(file);
						// renderPDF(Pdfpath);
						// extract(Pdfpath);
					} else {
						JOptionPane.showMessageDialog(textField_8, "Please select a PDF File");
					}
				} else {
					JOptionPane.showMessageDialog(textField_8, "Invalid Path");
				}

			}
		});

		btnNewButton_3.setBounds(503, 101, 78, 23);
		panel.add(btnNewButton_3);

		JLabel lblNewLabel_10 = new JLabel("Select the PDF File");
		lblNewLabel_10.setBounds(25, 77, 131, 14);
		panel.add(lblNewLabel_10);

		textField_4 = new JTextField();
		textField_4.setBounds(25, 326, 552, 20);
		panel.add(textField_4);
		textField_4.setColumns(10);

		JButton btnNewButton_5 = new JButton("Save Changes");
		btnNewButton_5.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (table.getSelectedRow() != -1) {
					String updatedData1 = textField.getText();
					String updatedData2 = textField_1.getText();
					String updatedData3 = textField_2.getText();
					String updatedData4 = textField_4.getText();
					String updatedData5 = textField_5.getText();
					String updatedData6 = textField_3.getText();
					String updatedData7 = textPane_JobDescription.getText();

					int selectedRow = table.getSelectedRow();
					tableModel.setValueAt(updatedData1, selectedRow, 0);
					tableModel.setValueAt(updatedData2, selectedRow, 1);
					tableModel.setValueAt(updatedData3, selectedRow, 2);
					tableModel.setValueAt(updatedData4, selectedRow, 3);
					tableModel.setValueAt(updatedData5, selectedRow, 4);
					tableModel.setValueAt(updatedData6, selectedRow, 5);
					tableModel.setValueAt(updatedData7, selectedRow, 6);

					String filePath = textField_7.getText();
					writeToExcel(filePath, tableModel);

					JOptionPane.showMessageDialog(frame, "Data Saved Successfully");
					textField.setText("");
					textField_1.setText("");
					textField_2.setText("");
					textField_4.setText("");
					textField_5.setText("");
					textField_3.setText("");
					textPane_JobDescription.setText("");
					textField_6.setText("");
					table.clearSelection();
				} else {
					JOptionPane.showMessageDialog(frame, "Please Select the Row From Table");
				}
			}
		});

		btnNewButton_5.setBounds(25, 877, 125, 23);
		panel.add(btnNewButton_5);
		table = new JTable(tableModel);
		JScrollPane tableScrollPane = new JScrollPane(table);
		tableScrollPane.setBounds(25, 636, 595, 230);
		panel.add(tableScrollPane);
		table.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
			public void valueChanged(ListSelectionEvent event) {
				if (!event.getValueIsAdjusting() && table.getSelectedRow() != -1) {

					int selectedRow = table.getSelectedRow();
					String data1 = table.getValueAt(selectedRow, 0).toString();
					String data2 = table.getValueAt(selectedRow, 1).toString();
					String data3 = table.getValueAt(selectedRow, 2).toString();
					String data4 = table.getValueAt(selectedRow, 3).toString();
					String data5 = table.getValueAt(selectedRow, 4).toString();
					String data6 = table.getValueAt(selectedRow, 5).toString();
					String data7 = table.getValueAt(selectedRow, 6).toString();

					textField.setText(data1);
					textField_1.setText(data2);
					textField_2.setText(data3);
					textField_4.setText(data4);
					textField_5.setText(data5);
					textField_3.setText(data6);
					textPane_JobDescription.setText(data7);
					int currentRowNumber = selectedRow + 1;
					totalRowCount = table.getRowCount();
					textField_6.setText(currentRowNumber + "/" + totalRowCount);
				}
			}
		});

	}

///////Save
	public static void writeToExcel(String filePath, DefaultTableModel tableModel) {
		try (FileInputStream inputStream = new FileInputStream(filePath);
				XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
				FileOutputStream outputStream = new FileOutputStream(filePath)) {

			XSSFSheet sheet = workbook.getSheetAt(0);
			int rowCount = tableModel.getRowCount();

			for (int i = 0; i < rowCount; i++) {
				XSSFRow row = sheet.getRow(i + 1);

				for (int j = 0; j < tableModel.getColumnCount(); j++) {
					Object value = tableModel.getValueAt(i, j);
					if (value != null) {
						XSSFCell cell = row.createCell(j);
						cell.setCellValue(value.toString());
					}
				}
			}
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

////////Excel
	public static void writeToExcel(String filePath, String data1, String data2, String data3, String data4,
			String data5, String data6, String data7, StyledDocument doc) {
		try (FileInputStream inputStream = new FileInputStream(filePath);
				XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
				FileOutputStream outputStream = new FileOutputStream(filePath)) {

			XSSFSheet sheet = workbook.getSheetAt(0);
			int rowCount = sheet.getLastRowNum();
			int emptyRowIndex = rowCount + 1;
			XSSFRow row = sheet.createRow(emptyRowIndex);

			XSSFCell cell1 = row.createCell(0);
			XSSFCell cell2 = row.createCell(1);
			XSSFCell cell3 = row.createCell(2);
			XSSFCell cell4 = row.createCell(3);
			XSSFCell cell5 = row.createCell(4);
			XSSFCell cell6 = row.createCell(5);
			XSSFCell cell7 = row.createCell(6);

			cell1.setCellValue(data1.trim());
			cell2.setCellValue(data2.trim());
			cell3.setCellValue(data3.trim());
			cell4.setCellValue(data4.trim());
			cell5.setCellValue(data5.trim());
			cell6.setCellValue(data6.trim());

			StringBuilder htmlBuilder = new StringBuilder();

			int start = 0;
			while (start < data7.length()) {
				int end = findWordEnd(data7, start);

				String word = data7.substring(start, end);
				Element element = doc.getCharacterElement(start);
				AttributeSet as = element.getAttributes();

				boolean isBold = StyleConstants.isBold(as);
				boolean isItalic = StyleConstants.isItalic(as);
				boolean isUnderline = StyleConstants.isUnderline(as);
				boolean isH1 = StyleConstants.getFontSize(as) == 24;
				boolean isH2 = StyleConstants.getFontSize(as) == 20;
				boolean isH3 = StyleConstants.getFontSize(as) == 18;
				boolean isH4 = StyleConstants.getFontSize(as) == 16;

				if (isBold)
					htmlBuilder.append("<p>" + "<b>");
				if (isItalic)
					htmlBuilder.append("<p>" + "<i>");
				if (isUnderline)
					htmlBuilder.append("<p>" + "<u>");
				if (isH1)
					htmlBuilder.append("<h1>");
				if (isH2)
					htmlBuilder.append("<h2>");
				if (isH3)
					htmlBuilder.append("<h3>");
				if (isH4)
					htmlBuilder.append("<h4>");
				if (!isBold && !isItalic && !isH1 && !isH2 && !isH3 && !isH4 && !isUnderline)
					htmlBuilder.append("<p>");
				htmlBuilder.append(word);
				if (isH4)
					htmlBuilder.append("</h4>");
				if (isH3)
					htmlBuilder.append("</h3>");
				if (isH2)
					htmlBuilder.append("</h2>");
				if (isH1)
					htmlBuilder.append("</h1>");

				if (isUnderline)
					htmlBuilder.append("</u>" + "</p>");
				if (isItalic)
					htmlBuilder.append("</i>" + "</p>");
				if (isBold)
					htmlBuilder.append("</b>" + "</p>");
				if (!isBold && !isItalic && !isH1 && !isH2 && !isH3 && !isH4 && !isUnderline)
					htmlBuilder.append("</p>");
				start = end;
			}

			cell7.setCellValue(htmlBuilder.toString().replace("\r\n" + "\r\n" + "<", ""));
			for (int i = 0; i <= rowCount; i++) {
				XSSFRow currentRow = sheet.getRow(i);
				if (currentRow != null) {
					for (int j = 0; j < currentRow.getLastCellNum(); j++) {
						XSSFCell cell = currentRow.getCell(j);
						if (cell != null && cell.getCellType() == CellType.STRING) {
							String cellValue = cell.getStringCellValue().trim();
							cell.setCellValue(cellValue);
						}
					}
				}
			}

			workbook.write(outputStream);

			System.out.println("Processsssssss Done");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static int findWordEnd(String text, int start) {
		int end = start;
		while (end < text.length() && Character.isLetterOrDigit(text.charAt(end))) {
			end++;
		}
		while (end < text.length() && !Character.isLetterOrDigit(text.charAt(end))) {
			end++;
		}
		return end;
	}

	private void previewExcelFile(String filePath) {
		try (FileInputStream file = new FileInputStream(filePath); XSSFWorkbook workbook = new XSSFWorkbook(file)) {
			XSSFSheet sheet = workbook.getSheetAt(0);
			tableModel.setRowCount(0);
			tableModel.setColumnCount(0);

			XSSFRow headerRow = sheet.getRow(0);
			if (headerRow != null) {
				for (int col = 0; col < headerRow.getPhysicalNumberOfCells(); col++) {
					XSSFCell cell = headerRow.getCell(col);
					tableModel.addColumn(cell.toString());
				}
			}
			for (int row = 1; row <= sheet.getLastRowNum(); row++) {
				XSSFRow excelRow = sheet.getRow(row);
				if (excelRow != null) {
					Object[] rowData = new Object[excelRow.getPhysicalNumberOfCells()];
					for (int col = 0; col < excelRow.getPhysicalNumberOfCells(); col++) {
						XSSFCell cell = excelRow.getCell(col);
						if (cell != null) {
							if (cell.getCellType() == CellType.NUMERIC) {
								rowData[col] = cell.getNumericCellValue();
							} else if (cell.getCellType() == CellType.BOOLEAN) {
								rowData[col] = cell.getBooleanCellValue();
							} else {
								rowData[col] = cell.getStringCellValue();
							}
						}
					}
					tableModel.addRow(rowData);
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(frame, "Please select the excel File!!!!");
		}
	}
}