package tool.Excel_to_XML;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class test {
public static void main (String excelFilePath) throws InvalidFormatException, IOException { 
    //public static void main(String[] args) throws IOException, InvalidFormatException {
        //String excelFilePath = "D:\\Template.xlsx"; 
        //String xmlFilePath = "D:\\output.xml";

        Workbook workbook = new XSSFWorkbook(new File(excelFilePath));
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row headerRow = rowIterator.next();
        StringBuilder xmlBuilder = new StringBuilder();
        xmlBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
        xmlBuilder.append("<root>\n");

        int rowNum = 1;
        while (rowIterator.hasNext()) {
            Row dataRow = rowIterator.next();
            xmlBuilder.append("  <row").append(rowNum).append(">\n");

            Iterator<Cell> cellIterator = headerRow.cellIterator();
            int cellIndex = 0;
            while (cellIterator.hasNext()) {
                Cell headerCell = cellIterator.next();
                String headerValue = headerCell.getStringCellValue().trim();
                Cell dataCell = dataRow.getCell(cellIndex);
                String dataValue = "";
                if (dataCell != null) {
                    switch (dataCell.getCellType()) {
                        case STRING:
                            dataValue = dataCell.getStringCellValue();
                            break;
                        case NUMERIC:
                            dataValue = String.valueOf(dataCell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            dataValue = String.valueOf(dataCell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            dataValue = dataCell.getCellFormula();
                            break;
                        default:
                            dataValue = "";
                    }
                }
                xmlBuilder.append("    <").append(headerValue).append(">").append(dataValue).append("</").append(headerValue).append(">\n");
                cellIndex++;
            }

            xmlBuilder.append("  </row").append(rowNum).append(">\n");
            rowNum++; 
        }
        xmlBuilder.append("</root>");
        try (FileOutputStream outputStream = new FileOutputStream("Output.xml")) {
            outputStream.write(xmlBuilder.toString().getBytes());
        }
        workbook.close();
        System.out.println("Process Done");
    }
}
