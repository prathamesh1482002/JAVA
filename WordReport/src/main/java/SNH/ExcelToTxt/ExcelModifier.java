package SNH.ExcelToTxt;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.*;
import java.util.*;
import javax.swing.JOptionPane;


public class ExcelModifier {
	public static void mainmethod(File file1, File file2, String path) {
		// public static void main(String[] args) {
		try {
			// FileInputStream file1 = new FileInputStream("D:\\prathmesh\\JAVA
			// Projects\\Label Request Ticket w PW - W Spine Bulk.xlsm");
			try (XSSFWorkbook workbook1 = new XSSFWorkbook(file1)) {
				XSSFSheet sheet1 = workbook1.getSheet("Label Request");
				// XSSFSheet sheet4 = workbook1.getSheet("Label Request_2");

				// FileInputStream file2 = new FileInputStream("D:\\prathmesh\\JAVA
				// Projects\\SET UP SHEET.xlsx");
				try (XSSFWorkbook workbook2 = new XSSFWorkbook(file2)) {
					XSSFSheet sheet2 = workbook2.getSheet("New");
					try (XSSFWorkbook newWorkbook = new XSSFWorkbook()) {
						XSSFSheet outputSheet = newWorkbook.createSheet("Output");
						// XSSFSheet outputSheet2 = newWorkbook.createSheet("QTY");

						CellStyle redStyle = newWorkbook.createCellStyle();
						Font redFont = newWorkbook.createFont();
						redFont.setColor(IndexedColors.RED.getIndex());
						redStyle.setFont(redFont);
						// String path = "D:\\prathmesh\\JAVA Projects\\Inputs";
						// File directory = new File(path);
						// Create a cell style
						CellStyle style = newWorkbook.createCellStyle();
						style.setAlignment(HorizontalAlignment.RIGHT);
						style.setVerticalAlignment(VerticalAlignment.BOTTOM);
						// BoLD
						CellStyle boldStyle = newWorkbook.createCellStyle();
						Font boldFont = newWorkbook.createFont();
						boldFont.setBold(true);
						boldStyle.setFont(boldFont);
						Row RC1 = sheet2.getRow(0); // 0-based index
						Cell CC1 = RC1.getCell(2);// B represents 1 in Apache POI
						String CClvalue = CC1.getStringCellValue();
						Row RH3 = sheet1.getRow(2);
						Cell CH3 = RH3.getCell(7);
						double CH3value = CH3.getNumericCellValue();
						// Job Name
						Row RA1 = sheet2.getRow(0);
						Cell CA1 = RA1.getCell(0);
//						// Job ID
						Row RG3 = sheet1.getRow(2);
						Cell CG3 = RG3.getCell(6);
						// volume /issue
						Row rowB4 = sheet1.getRow(3);
						Cell cellB4 = rowB4.getCell(1);

						Row rowC4 = sheet1.getRow(3);
						Cell cellC4 = rowC4.getCell(2);
						// AM
						Row rowE3 = sheet1.getRow(2);
						Cell CE3 = rowE3.getCell(4);

						Row RF3 = sheet1.getRow(2);
						Cell CF3 = RF3.getCell(5);
						// MAIl.Data
						Row A7 = sheet2.getRow(6);
						Cell CA7 = A7.getCell(0);

						String cellValue = CA7.getStringCellValue();
						String[] parts = cellValue.split(":", 2); // Split the cell value into two parts using ":"
						String secondPart = (parts.length > 1) ? parts[1].trim() : "";

						// MID (mailer ID):
						Row RH4 = sheet2.getRow(3);
						Cell CH4 = RH4.getCell(7);
						String MIDvalue = CH4.getStringCellValue();
						String[] MIDparts = MIDvalue.split(" ");
						String secnPart = MIDparts[1];
						// QTY
						// NCOA OCR
						Row ri = sheet1.getRow(20);
						Cell CB21 = ri.getCell(1);

						Row ry = sheet1.getRow(20);
						Cell CF21 = ry.getCell(5);
						// Postal
						Row b = sheet1.getRow(29);
						Cell CB30 = b.getCell(1);

						Row c = sheet1.getRow(29);
						Cell CF30 = c.getCell(5);
						// ADV
						Row d = sheet1.getRow(30);
						Cell CB31 = d.getCell(1);

						Row e = sheet1.getRow(30);
						Cell CF31 = e.getCell(5);
						// USPS Pub Number
						Row f = sheet1.getRow(31);
						Cell CB32 = f.getCell(1);

						Row g = sheet1.getRow(31);
						Cell CF32 = g.getCell(5);
						// Lable
						Row rhj = sheet1.getRow(22);
						Cell CB23 = rhj.getCell(1);

						Row rm = sheet1.getRow(22);
						Cell CF23 = rm.getCell(5);
						// Split
						Row RC7 = sheet1.getRow(6);
						Cell CC7 = RC7.getCell(2);

						Row RC8 = sheet1.getRow(7);
						Cell CC8 = RC8.getCell(2);
						// Piece Weight
						Row RD7 = sheet1.getRow(6);
						Cell CD7 = RD7.getCell(3);

						Row RD8 = sheet1.getRow(7);
						Cell CD8 = RD8.getCell(3);
						// Piece Thick
						Row RE7 = sheet1.getRow(6);
						Cell CE7 = RE7.getCell(4);

						Row RE8 = sheet1.getRow(7);
						Cell CE8 = RE8.getCell(4);
						// version
						Row RB7 = sheet1.getRow(6);
						Cell CB7 = RB7.getCell(1);

						Row RB8 = sheet1.getRow(7);
						Cell CB8 = RB8.getCell(1);
						// Trim
						Row rr = sheet1.getRow(15);
						Cell CB16 = rr.getCell(1);

						Row rrn = sheet1.getRow(15);
						Cell CF16 = rrn.getCell(5);
						// Permit
						Row ra = sheet1.getRow(16);
						Cell CB17 = ra.getCell(1);

						Row rb = sheet1.getRow(16);
						Cell CF17 = rb.getCell(5);
						// additional Instruction
						Row RA23 = sheet2.getRow(22);
						Cell CA23 = RA23.getCell(0);

						Row RA25 = sheet2.getRow(24);
						Cell CA25 = RA25.getCell(0);

						Row RA27 = sheet2.getRow(26);
						Cell CA27 = RA27.getCell(0);

						Row RA29 = sheet2.getRow(28);
						Cell CA29 = RA29.getCell(0);

						Row RA30 = sheet2.getRow(29);
						Cell CA30 = RA30.getCell(0);

						Row RA31 = sheet2.getRow(30);
						Cell CA31 = RA31.getCell(0);

						Row RA33 = sheet2.getRow(32);
						Cell CA33 = RA33.getCell(0);

						Row RA35 = sheet2.getRow(34);
						Cell CA35 = RA35.getCell(0);

						Row RA37 = sheet2.getRow(36);
						Cell CA37 = RA37.getCell(0);

						Row RA39 = sheet2.getRow(38);
						Cell CA39 = RA39.getCell(0);

						Row RA41 = sheet2.getRow(40);
						Cell CA41 = RA41.getCell(0);

//////////////////////////////// Write Start///////////////////////////////
						// Mail Preparation
						Row RNA1 = outputSheet.createRow(0);
						Cell CNA1 = RNA1.createCell(0);
						CNA1.setCellValue("Mail Preparation");
						//
						Row R2 = outputSheet.createRow(1);
						Cell newCell1 = R2.createCell(0);
						newCell1.setCellValue(CA1.getStringCellValue());

						Cell newCell2 = R2.createCell(1);
						newCell2.setCellValue(CC1.getStringCellValue());

						Row R4 = outputSheet.createRow(3);
						Cell CJD = R4.createCell(0);
						CJD.setCellValue(CG3.getStringCellValue());

						Cell CJ = R4.createCell(1);
						CJ.setCellValue(CH3.getNumericCellValue());
						//
						Row R6 = outputSheet.createRow(5);
						Cell AM = R6.createCell(0);
						AM.setCellValue(CE3.getStringCellValue());

						Cell MA = R6.createCell(1);
						MA.setCellValue(CF3.getStringCellValue());
						//
						Row R7 = outputSheet.createRow(6);
						Cell ML = R7.createCell(0);
						ML.setCellValue("MAIL.DAT FILENAME:");

						Cell ML2 = R7.createCell(1);
						ML2.setCellValue(secondPart);
						// MID
						Row R8 = outputSheet.createRow(7);
						Cell MID = R8.createCell(0);
						MID.setCellValue("MID (mailer ID):");

						Cell IDM = R8.createCell(1);
						IDM.setCellValue(secnPart);
						// Billl Inf
						Row R9 = outputSheet.createRow(8);
						Cell Bill = R9.createCell(0);
						Bill.setCellValue("Billing Info:");
						// QTY
						Row R00 = outputSheet.createRow(9);
						Cell qc = R00.createCell(0);
						qc.setCellValue("QTY");

						Cell OCc = R00.createCell(1);
						OCc.setCellValue("");

						// NCOA OCR
						Row R10 = outputSheet.createRow(11);
						Cell NC = R10.createCell(0);
						NC.setCellValue(CB21.getStringCellValue());

						Cell OC = R10.createCell(1);
						OC.setCellValue(CF21.getStringCellValue());
						// Postal
						Row R11 = outputSheet.createRow(12);
						Cell PO = R11.createCell(0);
						PO.setCellValue(CB30.getStringCellValue());

						Cell OP = R11.createCell(1);
						OP.setCellValue(CF30.getStringCellValue());
						// ADV
						Row R12 = outputSheet.createRow(13);
						Cell AD = R12.createCell(0);
						AD.setCellValue(CB31.getStringCellValue());

						Cell DA = R12.createCell(1);
						DA.setCellValue(CF31.getNumericCellValue());
						// USPS Pub Number
						Row R13 = outputSheet.createRow(14);
						Cell PUB = R13.createCell(0);
						PUB.setCellValue(CB32.getStringCellValue());

						Cell DUB = R13.createCell(1);
						DUB.setCellValue(CF32.getNumericCellValue());
						// Lable
						Row R14 = outputSheet.createRow(15);
						Cell LAB = R14.createCell(0);
						LAB.setCellValue(CB23.getStringCellValue());

						Cell ABS = R14.createCell(1);
						ABS.setCellValue(CF23.getStringCellValue());
						// Split
						Row R15 = outputSheet.createRow(16);
						Cell sp = R15.createCell(0);
						sp.setCellValue(CC7.getStringCellValue());

						Cell li = R15.createCell(1);
						li.setCellValue(CC8.getNumericCellValue());
						// Piece Weight
						Row R16 = outputSheet.createRow(17);
						Cell pi = R16.createCell(0);
						pi.setCellValue(CD7.getStringCellValue());

						Cell ec = R16.createCell(1);
						ec.setCellValue(CD8.getNumericCellValue());
						// piece Thick
						Row R17 = outputSheet.createRow(18);
						Cell th = R17.createCell(0);
						th.setCellValue(CE7.getStringCellValue());

						Cell ik = R17.createCell(1);
						ik.setCellValue(CE8.getNumericCellValue());
						// version
						Row R18 = outputSheet.createRow(19);
						Cell vv = R18.createCell(0);
						vv.setCellValue(CB7.getStringCellValue());

						Cell vr = R18.createCell(1);
						vr.setCellValue(CB8.getStringCellValue());
						// Trim
						Row R19 = outputSheet.createRow(20);
						Cell tr = R19.createCell(0);
						tr.setCellValue(CB16.getStringCellValue());

						Cell ti = R19.createCell(1);
						ti.setCellValue(CF16.getStringCellValue());
						// Permit
						Row R20 = outputSheet.createRow(21);
						Cell pr = R20.createCell(0);
						pr.setCellValue(CB17.getStringCellValue());

						Cell pit = R20.createCell(1);
						pit.setCellValue(CF17.getStringCellValue());
						// Job Issue
						Row R21 = outputSheet.createRow(4);
						Cell vi = R21.createCell(0);
						vi.setCellValue(cellB4.getStringCellValue());

						Cell ji = R21.createCell(1);
						ji.setCellValue(cellC4.getStringCellValue());
						// Abbreviation
						Row R22 = outputSheet.createRow(2);
						Cell abb = R22.createCell(0);
						abb.setCellValue("Abbreviation:");

						Cell ac = R22.createCell(1);
						ac.setCellValue("");
						// predrop Ship
						Row R23 = outputSheet.createRow(22);
						Cell pal = R23.createCell(0);
						pal.setCellValue("PRESORTDROP Ship:");

						Cell plll = R23.createCell(1);
						plll.setCellValue("");
						// pallets
						Row R24 = outputSheet.createRow(23);
						Cell pale = R24.createCell(0);
						pale.setCellValue("PALLETS:");

						Cell palet = R24.createCell(1);
						palet.setCellValue("");
						// TUB
						Row R25 = outputSheet.createRow(24);
						Cell Tub = R25.createCell(0);
						Tub.setCellValue("TUB:");

						Cell tube = R25.createCell(1);
						tube.setCellValue("");
						// Instructions
						Row R26 = outputSheet.createRow(25);
						Cell in = R26.createCell(0);
						in.setCellValue(CA23.getStringCellValue());

						Row R27 = outputSheet.createRow(26);
						Cell in2 = R27.createCell(0);
						in2.setCellValue(CA25.getStringCellValue());

						Row R29 = outputSheet.createRow(27);
						Cell in3 = R29.createCell(0);
						in3.setCellValue(CA27.getStringCellValue());

						Row R30 = outputSheet.createRow(28);
						Cell in4 = R30.createCell(0);
						in4.setCellValue(CA29.getStringCellValue());

						Row R31 = outputSheet.createRow(29);
						Cell in5 = R31.createCell(0);
						in5.setCellValue(CA30.getStringCellValue());

						Row R32 = outputSheet.createRow(30);
						Cell in6 = R32.createCell(0);
						in6.setCellValue(CA31.getStringCellValue());

						Row R33 = outputSheet.createRow(31);
						Cell in7 = R33.createCell(0);
						in7.setCellValue(CA33.getStringCellValue());

						Row R34 = outputSheet.createRow(32);
						Cell in8 = R34.createCell(0);
						in8.setCellValue(CA35.getStringCellValue());

						Row R35 = outputSheet.createRow(33);
						Cell in9 = R35.createCell(0);
						in9.setCellValue(CA37.getStringCellValue());

						Row R36 = outputSheet.createRow(34);
						Cell in10 = R36.createCell(0);
						in10.setCellValue(CA39.getStringCellValue());

						Row R37 = outputSheet.createRow(35);
						Cell in11 = R37.createCell(0);
						in11.setCellValue(CA41.getStringCellValue());

						// BOLD
						MID.setCellStyle(boldStyle);
						CNA1.setCellStyle(boldStyle);
						newCell1.setCellStyle(boldStyle);
						CJD.setCellStyle(boldStyle);
						AM.setCellStyle(boldStyle);
						ML.setCellStyle(boldStyle);
						Bill.setCellStyle(boldStyle);
						qc.setCellStyle(boldStyle);
						NC.setCellStyle(boldStyle);
						in.setCellStyle(boldStyle);
						Tub.setCellStyle(boldStyle);
						pale.setCellStyle(boldStyle);
						abb.setCellStyle(boldStyle);
						pal.setCellStyle(boldStyle);
						vi.setCellStyle(boldStyle);
						pr.setCellStyle(boldStyle);
						tr.setCellStyle(boldStyle);
						vv.setCellStyle(boldStyle);
						th.setCellStyle(boldStyle);
						pi.setCellStyle(boldStyle);
						sp.setCellStyle(boldStyle);
						LAB.setCellStyle(boldStyle);
						PUB.setCellStyle(boldStyle);
						AD.setCellStyle(boldStyle);
						PO.setCellStyle(boldStyle);
						// RED FONT
						in2.setCellStyle(redStyle);
						in3.setCellStyle(redStyle);
						in4.setCellStyle(redStyle);
						in5.setCellStyle(redStyle);
						in6.setCellStyle(redStyle);
						in7.setCellStyle(redStyle);
						in8.setCellStyle(redStyle);
						in9.setCellStyle(redStyle);
						in10.setCellStyle(redStyle);
						in11.setCellStyle(redStyle);
						// Text Formatting
						newCell2.setCellStyle(style);
						CJ.setCellStyle(style);
						MA.setCellStyle(style);
						IDM.setCellStyle(style);
						OCc.setCellStyle(style);
						OC.setCellStyle(style);
						OP.setCellStyle(style);
						DA.setCellStyle(style);
						DUB.setCellStyle(style);
						ABS.setCellStyle(style);
						li.setCellStyle(style);
						ec.setCellStyle(style);
						ik.setCellStyle(style);
						vr.setCellStyle(style);
						ti.setCellStyle(style);
						pit.setCellStyle(style);
						ji.setCellStyle(style);
						ac.setCellStyle(style);
						plll.setCellStyle(style);
						palet.setCellStyle(style);
						tube.setCellStyle(style);
						ML2.setCellStyle(style);
						// Write the new Excel file
						FileOutputStream outFile2 = new FileOutputStream(
								"Mail Prep_" + CClvalue + "_" + CH3value + ".xlsx");
						newWorkbook.write(outFile2);
						// Report writer Method
						System.out.println("Output Excel file has been created successfully.");
						ReportWriter(file1, file2, path);

						QTYWriter(file1, file2, path);
						System.out.println("QTY Writer Done.");

					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void ReportWriter(File file1, File file2, String path) throws InvalidFormatException, IOException {

		// String path = "D:\\prathmesh\\JAVA Projects\\Inputs";
		XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
		XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
		XSSFSheet sheet1 = workbook1.getSheet("Label Request");
		XSSFSheet sheet2 = workbook2.getSheet("New");

		Row RC1 = sheet2.getRow(0); 
		Cell CC1 = RC1.getCell(2);
		String CClvalue = CC1.getStringCellValue();
		Row RH3 = sheet1.getRow(2);
		Cell CH3 = RH3.getCell(7);
		double CH3value = CH3.getNumericCellValue();

		String Outputpath = "Mail Prep_" + CClvalue + "_" + CH3value + ".xlsx";
		String[] expectedHeaders = { "Member No", "Name", "C/O", "Address", "City", "State", "Zip", "Country",
				"Caption", "Status" };

		try {
			File reportFile = new File(Outputpath);
			XSSFWorkbook reportWorkbook;

			if (reportFile.exists()) {
				try (FileInputStream reportInputStream = new FileInputStream(reportFile)) {

					reportWorkbook = new XSSFWorkbook(reportInputStream);
				}
			} else {
				reportWorkbook = new XSSFWorkbook();
			}
			XSSFSheet reportSheet = reportWorkbook.getSheet("Report");
			if (reportSheet == null) {
				reportSheet = reportWorkbook.createSheet("Input Report");

				CellStyle boldStyle = reportWorkbook.createCellStyle();
				Font boldFont = reportWorkbook.createFont();
				boldFont.setBold(true);
				boldStyle.setFont(boldFont);

				Row headerRow = reportSheet.createRow(0);

				Cell fileNameHeader = headerRow.createCell(0);
				fileNameHeader.setCellValue("File Name");
				fileNameHeader.setCellStyle(boldStyle);

				Cell missingHeadersHeader = headerRow.createCell(1);
				missingHeadersHeader.setCellValue("Missing Headers");
				missingHeadersHeader.setCellStyle(boldStyle);

				Cell extraHeadersHeader = headerRow.createCell(2);
				extraHeadersHeader.setCellValue("Extra Headers");
				extraHeadersHeader.setCellStyle(boldStyle);
			}
			File inputDirectory = new File(path);
			File[] excelFiles = inputDirectory.listFiles(
					(dir, name) -> name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xlsm"));

			if (excelFiles != null) {
				int rowNum = reportSheet.getLastRowNum() + 1;

				for (File file : excelFiles) {
					try (FileInputStream fileInputStream = new FileInputStream(file);
							XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {
						XSSFSheet sheet = workbook.getSheetAt(0);
						int totalMissingHeadersCount = 0; 
						int totalExtraHeadersCount = 0;
						Row actualHeaderRow = sheet.getRow(0);
						List<String> actualHeaders = new ArrayList<>();
						for (Cell cell : actualHeaderRow) {
							actualHeaders.add(cell.getStringCellValue());
						}
						// Compare headers and find missing and extra headers
						List<String> missingHeaders = new ArrayList<>(Arrays.asList(expectedHeaders));
						missingHeaders.removeAll(actualHeaders);

						List<String> extraHeaders = new ArrayList<>(actualHeaders);
						extraHeaders.removeAll(Arrays.asList(expectedHeaders));
						// Add data to the report sheet
						Row dataRow = reportSheet.createRow(rowNum++);
						dataRow.createCell(0).setCellValue(file.getName());
						dataRow.createCell(1).setCellValue(String.join(", ", missingHeaders));
						dataRow.createCell(2).setCellValue(String.join(", ", extraHeaders));
						
						totalMissingHeadersCount += missingHeaders.size();
						totalExtraHeadersCount += extraHeaders.size();
						
						if (missingHeaders.size() > 2) {
							JOptionPane.showMessageDialog(null, + totalMissingHeadersCount +" Missing headers in " + file.getName(), "Missing Header Alert", JOptionPane.WARNING_MESSAGE);
				        }
						if (extraHeaders.size() != 0 ) {
							JOptionPane.showMessageDialog(null, + totalExtraHeadersCount +" Extra headers in " + file.getName(), "Extra Header Alert", JOptionPane.WARNING_MESSAGE);
				        }
	
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
				try (FileOutputStream outputStream = new FileOutputStream(Outputpath)) {
					reportWorkbook.write(outputStream);
					System.out.println("Input Report generated successfully.");
					reportWorkbook.close();
					workbook1.close();
					workbook2.close();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void QTYWriter(File file1, File file2, String path) {
		try (XSSFWorkbook workbook1 = new XSSFWorkbook(file1); XSSFWorkbook workbook2 = new XSSFWorkbook(file2)) {
			XSSFSheet sheet = workbook1.getSheet("Label Request");
			XSSFSheet sheet1 = workbook1.getSheet("Label Request_2");
			XSSFSheet sheet2 = workbook2.getSheet("New");

			Row RC1 = sheet2.getRow(0);
			Cell CC1 = RC1.getCell(2);
			String CClvalue = CC1.getStringCellValue();
			Row RH3 = sheet.getRow(2);
			Cell CH3 = RH3.getCell(7);
			double CH3value = CH3.getNumericCellValue();
			String Outputpath = "Mail Prep_" + CClvalue + "_" + CH3value + ".xlsx";
			File reportFile = new File(Outputpath);
			XSSFWorkbook reportWorkbook;

			if (reportFile.exists()) {
				try (FileInputStream reportInputStream = new FileInputStream(reportFile)) {
					reportWorkbook = new XSSFWorkbook(reportInputStream);
				}
			} else {
				reportWorkbook = new XSSFWorkbook();
			}
			XSSFSheet outputSheet = reportWorkbook.createSheet("QTY");

			// Ship#
			Row B66 = sheet1.getRow(65);
			Cell CB66 = B66.getCell(1);
			// Getter
			Row B67 = sheet1.getRow(66);
			Cell CB67 = B67.getCell(1);

			Row C67 = sheet1.getRow(66);
			Cell CC67 = C67.getCell(2);

			Row B68 = sheet1.getRow(67);
			Cell CB68 = B68.getCell(1);

			Row C68 = sheet1.getRow(67);
			Cell CC68 = C68.getCell(2);

			Row B69 = sheet1.getRow(68);
			Cell CB69 = B69.getCell(1);

			Row C69 = sheet1.getRow(68);
			Cell CC69 = C69.getCell(2);

			Row B70 = sheet1.getRow(69);
			Cell CB70 = B70.getCell(1);

			Row C70 = sheet1.getRow(69);
			Cell CC70 = C70.getCell(2);

			Row B71 = sheet1.getRow(70);
			Cell CB71 = B71.getCell(1);

			Row C71 = sheet1.getRow(70);
			Cell CC71 = C71.getCell(2);

			Row B72 = sheet1.getRow(71);
			Cell CB72 = B72.getCell(1);

			Row C72 = sheet1.getRow(71);
			Cell CC72 = C72.getCell(2);

			// Writer
			if (CB67.toString() != "") {
				Row R35 = outputSheet.createRow(34);
				Cell row36 = R35.createCell(0);
				row36.setCellValue(CB66 + " " + CB67.toString());

				Cell C35 = R35.createCell(10);
				C35.setCellValue(CC67.toString());
			}
			///
			if (CB68.toString() != "") {
				Row R36 = outputSheet.createRow(35);
				Cell row37 = R36.createCell(0);
				row37.setCellValue(CB66 + " " + CB68.toString());

				Cell C36 = R36.createCell(10);
				C36.setCellValue(CC68.toString());
			}
			if (CB69.toString() != "") {
				Row R37 = outputSheet.createRow(36);
				Cell row38 = R37.createCell(0);
				row38.setCellValue(CB66 + " " + CB69.toString());

				Cell C37 = R37.createCell(10);
				C37.setCellValue(CC69.toString());
			}
			if (CB70.toString() != "") {
				Row R38 = outputSheet.createRow(37);
				Cell row39 = R38.createCell(0);
				row39.setCellValue(CB66 + " " + CB70.toString());

				Cell C38 = R38.createCell(10);
				C38.setCellValue(CC70.toString());
			}
			if (CB71.toString() != "") {
				Row R39 = outputSheet.createRow(38);
				Cell row40 = R39.createCell(0);
				row40.setCellValue(CB66 + " " + CB71.toString());

				Cell C39 = R39.createCell(10);
				C39.setCellValue(CC71.toString());
			}
			if (CB72.toString() != "") {
				Row R40 = outputSheet.createRow(39);
				Cell row41 = R40.createCell(0);
				row41.setCellValue(CB66 + " " + CB72.toString());

				Cell C40 = R40.createCell(10);
				C40.setCellValue(CC72.toString());
			}
			// Create header row
			CellStyle boldStyle = reportWorkbook.createCellStyle();
			Font boldFont = reportWorkbook.createFont();
			boldFont.setBold(true);
			boldStyle.setFont(boldFont);

			Row headerRow1 = outputSheet.createRow(0);
			Cell headerCell1 = headerRow1.createCell(0);
			headerCell1.setCellValue("File Name");
			headerCell1.setCellStyle(boldStyle);
			Cell headerCell2 = headerRow1.createCell(1);
			headerCell2.setCellValue("Row Count");
			headerCell2.setCellStyle(boldStyle);

			int rowIndex = 1;
			File[] excelFiles = new File(path).listFiles();

			if (excelFiles != null) {
				for (File excelFile : excelFiles) {
					try (Workbook workbookInput = new XSSFWorkbook(new FileInputStream(excelFile))) {
						Sheet inputSheet = workbookInput.getSheetAt(0);
						int rowCount = inputSheet.getPhysicalNumberOfRows();
						// Write file name and row count to the "QTY" sheet
						Row dataRow = outputSheet.createRow(rowIndex++);
						Cell fileNameCell = dataRow.createCell(0);
						fileNameCell.setCellValue(excelFile.getName());
						Cell rowCountCell = dataRow.createCell(1);
						rowCountCell.setCellValue(rowCount - 1);
					} catch (IOException e) {
						e.printStackTrace(); 
					}
				}
			}

			int newRowNum = 17;

			for (int i = 0; i <= sheet1.getLastRowNum(); i++) {
				XSSFRow row = sheet1.getRow(i);

				if (row != null) {
					XSSFRow newRow = outputSheet.createRow(newRowNum++);
					int lastCellIndex = row.getLastCellNum();

					for (int j = 1; j < lastCellIndex; j++) { 
						XSSFCell cell = row.getCell(j);

						if (cell != null && !cell.toString().isEmpty()) {
							String cellValue = cell.toString();

							if (cellValue.equals("L/O=Paper Label on plastic")) {
								
								try (FileOutputStream outputStream = new FileOutputStream(Outputpath)) {
									reportWorkbook.write(outputStream);
									reportWorkbook.close();
									System.out.println("QTY sheet generated successfully.");
								} catch (IOException e) {
									e.printStackTrace();
								}
								return;
							}

							XSSFCell newCell = newRow.createCell(j - 1); 
							newCell.setCellValue(cellValue);

							
							if (newRowNum == 18 || newRowNum == 19) {
								newCell.setCellStyle(boldStyle);
							}
						}
					}
				}
			}
			reportWorkbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}