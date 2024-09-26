package XML_Invoicing.XML_Billing;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class XML_Invoicing {
	public static void main(String excelFilePath, String invoiceNumber) {
	//public static void main(String[] args) {
	    try {
	       // String excelFilePath = "D:\\Prathamesh\\Final Testing\\TAF Artwork Extraction-August-2024-Kerry.xls"; 
	        //String invoiceNumber = "KGL-10885-32";
	        String documentId = generateDocumentId(invoiceNumber);

	        Workbook workbook = getWorkbook(excelFilePath);
	        if (workbook == null) {
	            System.out.println("Unable to open workbook. Check file format.");
	            return;
	        }
	        String sheetName = workbook.getSheetName(1);
	        System.out.println("Processing sheet: " + sheetName); 
	        Sheet sheet = workbook.getSheet(sheetName);
	            double docValue = calculateDocValue(sheet); 
	            String xmlData = generateXML(sheet, invoiceNumber, documentId, docValue);

	            writeXMLToFile(xmlData, excelFilePath, documentId);
	            System.out.println("XML File created successfully.");
	      

	        workbook.close();
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}


    private static Workbook getWorkbook(String excelFilePath) throws IOException {
        Workbook workbook = null;
        try (FileInputStream fis = new FileInputStream(excelFilePath)) {
            if (excelFilePath.toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (excelFilePath.toLowerCase().endsWith(".xls")) {
                POIFSFileSystem poifsFileSystem = new POIFSFileSystem(fis);
                workbook = new HSSFWorkbook(poifsFileSystem); 
            }
        }
        return workbook;
    }

    private static double calculateDocValue(Sheet sheet) {
        double totalValue = 0.0;
        int lastColumnIndex = sheet.getRow(7).getLastCellNum() - 1;
        int lastRowNum = sheet.getLastRowNum();
        for (int i = lastRowNum; i >= 7; i--) {
            Row row = sheet.getRow(i);
            if (row != null && row.getCell(lastColumnIndex) != null) {
                String cellValue = getCellValueAsString(row.getCell(lastColumnIndex));
                if (!cellValue.isEmpty()) {
                    try {
                        totalValue = Double.parseDouble(cellValue); 
                        break;
                    } catch (NumberFormatException e) {
                        System.out.println("Invalid number format in row " + (i + 1));
                    }
                }
            }
        }
        return Math.round(totalValue * 100.0) / 100.0; 
    }

    private static String generateXML(Sheet sheet, String invoiceNumber, String documentId, double docValue) {
        StringBuilder xmlBuilder = new StringBuilder();

        xmlBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
        xmlBuilder.append("<edi xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n");
        xmlBuilder.append("<files>\n");
        xmlBuilder.append("<file type=\"INVOICES\">\n");


        xmlBuilder.append("<header>\n");
        xmlBuilder.append("<document>").append(documentId).append(".pdf</document>\n");
        xmlBuilder.append("<sourceID>SUPPLIER_XML</sourceID>\n");
        xmlBuilder.append("<supplierID>1000091394</supplierID>\n");
        xmlBuilder.append("<supplierName>KnowledgeWorks Global LTD</supplierName>\n");
        xmlBuilder.append("<description>").append(getCellValueAsString(sheet.getRow(1).getCell(1))).append("</description>\n");
        xmlBuilder.append("<invoiceNumber>").append(invoiceNumber).append("</invoiceNumber>\n");
        xmlBuilder.append("<curCode>GBP</curCode>\n");
        xmlBuilder.append("<invDate>").append(new SimpleDateFormat("dd-MM-yyyy").format(new Date())).append("</invDate>\n");
        xmlBuilder.append("<cmpCode>INF1</cmpCode>\n");
        xmlBuilder.append("<approverID>").append(getCellValueAsString(sheet.getRow(2).getCell(1))).append("</approverID>\n");
        xmlBuilder.append("<profitCentre>2GDGB17400</profitCentre>\n");
        xmlBuilder.append("<docValue>").append(String.valueOf(docValue)).append("</docValue>\n");
        xmlBuilder.append("<docTaxAmount></docTaxAmount>\n");
        xmlBuilder.append("<docType>Invoice</docType>\n");
        xmlBuilder.append("</header>\n");

        xmlBuilder.append("<lines>\n");

        for (int i = 6; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null || row.getCell(3) == null || getCellValueAsString(row.getCell(3)).isEmpty()) {
                break;
            }
            String origin = getCellValueAsString(row.getCell(3)).trim();
            if ("UK".equalsIgnoreCase(origin) || "US".equalsIgnoreCase(origin)) {
                xmlBuilder.append("<line>\n");
                xmlBuilder.append("<description>").append(getCellValueAsString(sheet.getRow(1).getCell(1))).append("</description>\n");
                xmlBuilder.append("<materialNumber>259900100</materialNumber>\n");

                String wbsCode = "BO." + getCellValueAsString(row.getCell(2)) + (origin.equals("UK") ? ".INF1.97" : ".INF3.97");
                xmlBuilder.append("<WBSCode>").append(wbsCode).append("</WBSCode>\n");
                xmlBuilder.append("<costCentre></costCentre>\n");
                xmlBuilder.append("<tradingPartner>").append(origin.equals("US") ? "IINF3" : "").append("</tradingPartner>\n");
                xmlBuilder.append("<fixedAsset></fixedAsset>\n");
                xmlBuilder.append("<internalOrder></internalOrder>\n");
                xmlBuilder.append("<lineProfitCentre>").append(origin.equals("US") ? "2GDGB17401" : "").append("</lineProfitCentre>\n");
                xmlBuilder.append("<manuscriptID></manuscriptID>\n");
                xmlBuilder.append("<lineTaxAmount></lineTaxAmount>\n");
                xmlBuilder.append("<lineTaxType>ZRO</lineTaxType>\n");
                xmlBuilder.append("<lineTaxPercent></lineTaxPercent>\n");

                int lastColumnIndex = row.getLastCellNum() - 1;
                String lastColumnValue = getCellValueAsString(row.getCell(lastColumnIndex));

                // If the last column is blank, take the value from the previous column
                if (lastColumnValue == null || lastColumnValue.trim().isEmpty()) {
                    lastColumnValue = getCellValueAsString(row.getCell(lastColumnIndex - 1));
                }

                xmlBuilder.append("<lineValue>").append(lastColumnValue).append("</lineValue>\n");
                xmlBuilder.append("<lineType>Debit</lineType>\n");
                xmlBuilder.append("</line>\n");
            }
        }


        xmlBuilder.append("</lines>\n");
        xmlBuilder.append("</file>\n");
        xmlBuilder.append("</files>\n");
        xmlBuilder.append("</edi>\n");

        return xmlBuilder.toString();
    }


    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                } else {
                    return String.format("%.2f", cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                switch (evaluator.evaluateFormulaCell(cell)) {
                    case NUMERIC:
                        return String.format("%.2f", evaluator.evaluate(cell).getNumberValue());
                    case STRING:
                        return evaluator.evaluate(cell).getStringValue();
                    case BOOLEAN:
                        return String.valueOf(evaluator.evaluate(cell).getBooleanValue());
                    default:
                        return "";
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }


    private static void writeXMLToFile(String xmlData, String excelFilePath, String documentId) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(new File(new File(excelFilePath).getParent(), documentId + ".xml"));
        outputStream.write(xmlData.getBytes());
        outputStream.close();
    }

    private static String generateDocumentId(String invoiceNumber) {
        String timestamp = new SimpleDateFormat("yyyyMMddHHmm").format(new Date());
        return "EDI_" + invoiceNumber + "_1000091394_" + timestamp;
    }
}
