package XML_Invoicing.XML_Billing;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.*;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Summary {

    public static void processXMLFiles(String directoryPath, String outputFileName) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet activeSheet = workbook.createSheet("Summary");

            Row headerRow = activeSheet.createRow(12);
            headerRow.createCell(10).setCellValue("Sr#");
            headerRow.createCell(11).setCellValue("Invoice#");
            headerRow.createCell(12).setCellValue("Value");
            headerRow.createCell(13).setCellValue("Approver");

            CellStyle boldStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            boldStyle.setFont(font);
            for (int i = 10; i <= 13; i++) {
                headerRow.getCell(i).setCellStyle(boldStyle);
            }
            File dir = new File(directoryPath);
            File[] files = dir.listFiles((d, name) -> name.endsWith(".xml"));

            if (files != null) {
                int rowNumber = 13; 
                int serialNumber = 1;

                for (File file : files) {
                    String fileName = file.getName();

                    String invoiceNumber = "";
                    if (fileName.contains("_")) {
                        String[] fileNameParts = fileName.split("_");

                        if (fileNameParts.length >= 2) {
                            invoiceNumber = fileNameParts[0] + "_" + fileNameParts[1];
                        } else {
                            System.err.println("Invalid filename format: " + fileName);
                        }
                    } else {
                        System.err.println("Invalid filename format: " + fileName);
                    }

                    String docValue = "";
                    String approverID = "";
                    try {
                        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                        DocumentBuilder builder = factory.newDocumentBuilder();
                        Document document = builder.parse(file);

                        NodeList docValueNodes = document.getElementsByTagName("docValue");
                        if (docValueNodes.getLength() > 0) {
                            docValue = docValueNodes.item(0).getTextContent();
                        }

                        NodeList approverIDNodes = document.getElementsByTagName("approverID");
                        if (approverIDNodes.getLength() > 0) {
                            approverID = approverIDNodes.item(0).getTextContent();
                        }

                    } catch (Exception e) {
                        System.err.println("Error parsing XML file: " + fileName);
                        e.printStackTrace();
                    }
                    Row row = activeSheet.createRow(rowNumber++);

                    row.createCell(10).setCellValue(serialNumber++);  
                    row.createCell(11).setCellValue(invoiceNumber);   
                    row.createCell(12).setCellValue(docValue);         
                    row.createCell(13).setCellValue(approverID);       
                }


            }

            try (FileOutputStream fileOut = new FileOutputStream(directoryPath+"//"+outputFileName+""+".xlsx")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                System.err.println("Error writing Excel file: " + e.getMessage());
                e.printStackTrace();
            }

        }
    }
}
