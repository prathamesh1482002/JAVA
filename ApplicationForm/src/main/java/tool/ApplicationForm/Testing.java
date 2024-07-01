package tool.ApplicationForm;


import java.awt.*;
import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.text.*;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;

import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.event.ActionEvent;

public class Testing {

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
	private JTextPane textPane_PDFData;
	private JTable table;
	private DefaultTableModel tableModel;



	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Testing window = new Testing();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public Testing() {
		initialize();
	}

	private void initialize() {
		frame = new JFrame();
		frame.setTitle("Application Form");
		frame.setBounds(100, 100, 1523, 765);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);

		textField = new JTextField();
		textField.setBounds(25, 102, 552, 20);
		frame.getContentPane().add(textField);
		textField.setColumns(10);

		textField_1 = new JTextField();
		textField_1.setBounds(25, 158, 552, 20);
		frame.getContentPane().add(textField_1);
		textField_1.setColumns(10);

		textField_2 = new JTextField();
		textField_2.setBounds(25, 214, 552, 20);
		frame.getContentPane().add(textField_2);
		textField_2.setColumns(10);

		JLabel lblNewLabel = new JLabel("Input 1");
		lblNewLabel.setBounds(25, 77, 61, 14);
		frame.getContentPane().add(lblNewLabel);

		JLabel lblNewLabel_1 = new JLabel("Input 2");
		lblNewLabel_1.setBounds(25, 133, 46, 14);
		frame.getContentPane().add(lblNewLabel_1);

		JLabel lblNewLabel_2 = new JLabel("Input 3");
		lblNewLabel_2.setBounds(25, 189, 61, 14);
		frame.getContentPane().add(lblNewLabel_2);

		JLabel lblNewLabel_3 = new JLabel("Input 4");
		lblNewLabel_3.setBounds(25, 245, 46, 14);
		frame.getContentPane().add(lblNewLabel_3);

		JLabel lblNewLabel_4 = new JLabel("Input 6");
		lblNewLabel_4.setBounds(25, 355, 86, 14);
		frame.getContentPane().add(lblNewLabel_4);

		textField_3 = new JTextField();
		textField_3.setBounds(25, 380, 552, 20);
		frame.getContentPane().add(textField_3);
		textField_3.setColumns(10);

		JLabel lblNewLabel_5 = new JLabel("Input 7");
		lblNewLabel_5.setBounds(25, 411, 109, 14);
		frame.getContentPane().add(lblNewLabel_5);
		
		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
		scrollPane_1.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane_1.setBounds(25, 436, 309, 100);
		frame.getContentPane().add(scrollPane_1);

		textPane_JobDescription = new JTextPane();
		scrollPane_1.setViewportView(textPane_JobDescription);

		JButton btnBold = new JButton("B");
		btnBold.setFont(new Font("Arial", Font.BOLD, 12));
		btnBold.setBounds(344, 436, 52, 25);
		frame.getContentPane().add(btnBold);

		JButton btnItalic = new JButton("I");
		btnItalic.setFont(new Font("Arial", Font.ITALIC, 12));
		btnItalic.setBounds(463, 436, 52, 25);
		frame.getContentPane().add(btnItalic);

		JButton btnUnderline = new JButton("U");
		btnUnderline.setFont(new Font("Arial", Font.PLAIN, 12));
		btnUnderline.setBounds(401, 436, 52, 25);
		frame.getContentPane().add(btnUnderline);

		JButton btnPlainText = new JButton("P");
		btnPlainText.setFont(new Font("Arial", Font.PLAIN, 12));
		btnPlainText.setBounds(525, 436, 52, 25);
		frame.getContentPane().add(btnPlainText);

		tableModel = new DefaultTableModel();
		table = new JTable(tableModel);
		JScrollPane tableScrollPane = new JScrollPane(table);
		tableScrollPane.setBounds(21, 593, 556, 120);
		frame.getContentPane().add(tableScrollPane);
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
					int totalRowCount = table.getRowCount();
					textField_6.setText(currentRowNumber + "/" + totalRowCount);
				}
			}
		});

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
		JButton btnNewButton = new JButton("Submit");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
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
			}
		});
		btnNewButton.setBounds(364, 904, 89, 23);
		frame.getContentPane().add(btnNewButton);

		JLabel lblNewLabel_7 = new JLabel("Input 5");
		lblNewLabel_7.setBounds(25, 301, 45, 13);
		frame.getContentPane().add(lblNewLabel_7);

		textField_5 = new JTextField();
		textField_5.setBounds(25, 325, 552, 19);
		frame.getContentPane().add(textField_5);
		textField_5.setColumns(10);

		JLabel lblNewLabel_9 = new JLabel("Select the Excel File");
		lblNewLabel_9.setBounds(25, 23, 125, 13);
		frame.getContentPane().add(lblNewLabel_9);

		textField_7 = new JTextField();
		textField_7.setBounds(25, 47, 447, 19);
		frame.getContentPane().add(textField_7);
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
		frame.getContentPane().add(btnNewButton_1);
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
		btnNewButton_2.setBounds(344, 513, 52, 23);
		frame.getContentPane().add(btnNewButton_2);

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
		btnNewButton_2_1.setBounds(401, 513, 54, 23);
		frame.getContentPane().add(btnNewButton_2_1);

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
		btnNewButton_2_1_1.setBounds(463, 513, 52, 23);
		frame.getContentPane().add(btnNewButton_2_1_1);

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
		btnNewButton_2_1_1_1.setBounds(525, 513, 52, 23);
		frame.getContentPane().add(btnNewButton_2_1_1_1);

		JLabel lblNewLabel_8 = new JLabel("Current RowNumber/Total Row Number");
		lblNewLabel_8.setBounds(25, 557, 227, 13);
		frame.getContentPane().add(lblNewLabel_8);

		textField_6 = new JTextField();
		textField_6.setBounds(262, 553, 72, 19);
		frame.getContentPane().add(textField_6);
		textField_6.setColumns(10);
		textField_8 = new JTextField();
		textField_8.setBounds(792, 47, 403, 20);
		frame.getContentPane().add(textField_8);
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
					String Pdfpath = textField_8.getText();
					renderPDF(Pdfpath);
					

				} else {
					JOptionPane.showMessageDialog(textField_8, "Invalid Path");
				}

			}
		});
		
		btnNewButton_3.setBounds(1213, 46, 78, 23);
		frame.getContentPane().add(btnNewButton_3);

		JLabel lblNewLabel_10 = new JLabel("Select the PDF File");
		lblNewLabel_10.setBounds(651, 50, 131, 14);
		frame.getContentPane().add(lblNewLabel_10);

		textField_4 = new JTextField();
		textField_4.setBounds(25, 270, 552, 20);
		frame.getContentPane().add(textField_4);
		textField_4.setColumns(10);

		JButton btnNewButton_5 = new JButton("Save");
		btnNewButton_5.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

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
				table.clearSelection();
				
			}
		});

		btnNewButton_5.setBounds(137, 904, 89, 23);
		frame.getContentPane().add(btnNewButton_5);
		
		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(663, 79, 836, 457);
		frame.getContentPane().add(scrollPane);
		
		JScrollPane scrollPane_2 = new JScrollPane();
		scrollPane_2.setBounds(663, 568, 836, 145);
		frame.getContentPane().add(scrollPane_2);
		
		textPane_JobDescription = new JTextPane();

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

			System.out.println("Process Done");
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
			JOptionPane.showMessageDialog(frame, "Please select the excel File!!!");
		}
	}
	private void renderPDF(String pdfPath) {
		 try {
		        PDDocument document = PDDocument.load(new File(pdfPath));
		        PDFRenderer renderer = new PDFRenderer(document);
		        int numPages = document.getNumberOfPages();
		        JPanel pdfPanel = new JPanel();
		        pdfPanel.setLayout(new BoxLayout(pdfPanel, BoxLayout.Y_AXIS));
		        JTextArea textArea = new JTextArea();
		        textArea.setEditable(false);
		        textArea.setLineWrap(true);
		        textArea.setWrapStyleWord(true);
		        
		        for (int i = 0; i < numPages; i++) {
		            BufferedImage img = renderer.renderImageWithDPI(i, 100);
		            JLabel imgLabel = new JLabel(new ImageIcon(img));
		            JTextArea textOverlay = new JTextArea();
		            textOverlay.setOpaque(false); // Make the text area transparent
		            textOverlay.setLineWrap(true);
		            textOverlay.setWrapStyleWord(true);

		            // Extract text for the page
		            PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfPath));
		            String pageText = PdfTextExtractor.getTextFromPage(pdfDoc.getPage(i + 1));
		            textOverlay.setText(pageText);
		            textOverlay.addMouseListener(new MouseAdapter() {
		                @Override
		                public void mousePressed(MouseEvent e) {
		                    textOverlay.requestFocus();
		                }
		            });
		            textOverlay.getDocument().addDocumentListener(new DocumentListener() {
		                @Override
		                public void insertUpdate(DocumentEvent e) {
		                    updateTextArea();
		                }

		                @Override
		                public void removeUpdate(DocumentEvent e) {
		                    updateTextArea();
		                }

		                @Override
		                public void changedUpdate(DocumentEvent e) {
		                    updateTextArea();
		                }

		                private void updateTextArea() {
		                    textArea.setText(textOverlay.getSelectedText());
		                }
		            });

		            JPanel pagePanel = new JPanel();
		            pagePanel.setLayout(new OverlayLayout(pagePanel));
		            pagePanel.add(textOverlay);
		            pagePanel.add(imgLabel);
		            pdfPanel.add(pagePanel);
		        }
		        
		        JScrollPane pdfScrollPane = new JScrollPane(pdfPanel);
		        pdfScrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
		        pdfScrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		        pdfScrollPane.setBounds(639, 77, 858, 636);
		        frame.getContentPane().add(pdfScrollPane);
		        
		        // Text area for displaying selected text
		        JScrollPane textScrollPane = new JScrollPane(textArea);
		        textScrollPane.setBounds(10, 77, 600, 636);
		        frame.getContentPane().add(textScrollPane);

		        pdfScrollPane.setVisible(true);
		        textScrollPane.setVisible(true);
		        frame.revalidate();
		        frame.repaint();
		        document.close();
		    } catch (Exception ex) {
		        ex.printStackTrace();
		        JOptionPane.showMessageDialog(frame, "Error " + ex.getMessage());
		    }
	}

	private void Reader(String Pdfpath) {
	    try {
	        PDDocument document = PDDocument.load(new File(Pdfpath));
	        PDFTextStripper pdfStripper = new PDFTextStripper();
	        String pdfText = pdfStripper.getText(document);

	        // Set the extracted PDF text to the textPane_PDFData
	        textPane_PDFData.setText(pdfText);

	        document.close();
	    } catch (Exception ex) {
	        ex.printStackTrace();
	        JOptionPane.showMessageDialog(frame, "Error " + ex.getMessage());
	    }
	}
}