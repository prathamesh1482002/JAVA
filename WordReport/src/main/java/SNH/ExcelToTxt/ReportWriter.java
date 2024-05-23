package SNH.ExcelToTxt;
import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.util.*;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReportWriter {
    private static String[] expectedHeaders = { "Member No", "Name", "C/O", "Address", "City", "State", "Zip", "Country", "Caption", "Status" };

    private static void addHeaderToSheet(XSSFSheet sheet, String extraHeader, String expectedHeader) {
        Row headerRow = sheet.getRow(0);
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null && extraHeader.equals(cell.getStringCellValue())) {
                cell.setCellValue(expectedHeader);
                break;
            }
        }
    }

    private static void reportMissingHeaders(File file, List<String> missingHeaders, XSSFWorkbook reportWorkbook, XSSFSheet reportSheet) {
        int rowNum = reportSheet.getLastRowNum() + 1;

        JPanel panel = new JPanel(new GridLayout(0, 2));
        JComboBox<String> extraHeadersDropdown = new JComboBox<>(missingHeaders.toArray(new String[0]));
        JComboBox<String> expectedHeadersDropdown = new JComboBox<>(expectedHeaders);
        panel.add(new JLabel("Select Extra Header:"));
        panel.add(extraHeadersDropdown);
        panel.add(new JLabel("Select Expected Header:"));
        panel.add(expectedHeadersDropdown);

        int result = JOptionPane.showConfirmDialog(null, panel, "Select Headers", JOptionPane.OK_CANCEL_OPTION);
        if (result == JOptionPane.OK_OPTION) {
            String extraHeader = (String) extraHeadersDropdown.getSelectedItem();
            String expectedHeader = (String) expectedHeadersDropdown.getSelectedItem();
            addHeaderToSheet(reportSheet, extraHeader, expectedHeader);
        }

        Row dataRow = reportSheet.createRow(rowNum);
        dataRow.createCell(0).setCellValue(file.getName());
        dataRow.createCell(1).setCellValue(String.join(", ", missingHeaders));
        dataRow.createCell(2).setCellValue(String.join(", ", (CharSequence[]) extraHeadersDropdown.getSelectedItem()));

        if (missingHeaders.size() > 2) {
            try (FileOutputStream outputStream = new FileOutputStream("Report.xlsx")) {
                reportWorkbook.write(outputStream);
                System.out.println("Input Report generated successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void analyzeAndReport(File file, XSSFWorkbook reportWorkbook, XSSFSheet reportSheet) {
        try (FileInputStream fileInputStream = new FileInputStream(file)) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Row actualHeaderRow = sheet.getRow(0);
            List<String> actualHeaders = new ArrayList<>();
            for (Cell cell : actualHeaderRow) {
                actualHeaders.add(cell.getStringCellValue());
            }

            List<String> missingHeaders = new ArrayList<>(Arrays.asList(expectedHeaders));
            missingHeaders.removeAll(actualHeaders);

            List<String> extraHeaders = new ArrayList<>(actualHeaders);
            extraHeaders.removeAll(Arrays.asList(expectedHeaders));

            if (missingHeaders.size() > 2) {
                reportMissingHeaders(file, missingHeaders, reportWorkbook, reportSheet);
            } else {
                int rowNum = reportSheet.getLastRowNum() + 1;
                Row dataRow = reportSheet.createRow(rowNum);
                dataRow.createCell(0).setCellValue(file.getName());
                dataRow.createCell(1).setCellValue(String.join(", ", missingHeaders));
                dataRow.createCell(2).setCellValue(String.join(", ", extraHeaders));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws InvalidFormatException {
        String path = "D:\\SNH\\Inputs";
        File file1 = new File("D:\\SNH\\Label Request Ticket w PW - W Spine Bulk.xlsm");
        File file2 = new File("D:\\SNH\\SET UP SHEET.xlsx");

        try (XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
             XSSFWorkbook workbook2 = new XSSFWorkbook(file2);
             XSSFWorkbook reportWorkbook = new XSSFWorkbook()) {

            XSSFSheet sheet1 = workbook1.getSheet("Label Request");
            XSSFSheet sheet2 = workbook2.getSheet("New");

            XSSFSheet reportSheet = reportWorkbook.createSheet("Report");

            analyzeAndReport(file1, reportWorkbook, reportSheet);
            analyzeAndReport(file2, reportWorkbook, reportSheet);

            try (FileOutputStream outputStream = new FileOutputStream("Report.xlsx")) {
                reportWorkbook.write(outputStream);
                System.out.println("Input Report generated successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
