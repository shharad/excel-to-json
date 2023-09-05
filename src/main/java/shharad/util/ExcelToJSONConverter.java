package shharad.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * @AUther: smeena
 * @Date: 05-Sep-2023 12:44:34
 *
 * This utility class reads the Data from Excel Sheet and converts it into JSON array.
 * It currently handles two levels of Header Row to generate equivalent nested JSON object.
 * You can enhance the code to support more 2 levels of header/sub-header rows.
 * For single header row also the code changes is required.
 *
 * Maven Version used: apache-maven-3.6.3
 *
 */
public class ExcelToJSONConverter {

    /**
     * Main method to test this converter
     *
     * @param args
     */
    public static void main(String[] args) {

        // Creating a file object with specific file path
        // File excel = new File("C:\\TEMP\\SampleData-ExcelToJSON.xlsx"); // location of excel file in the local file system
        File excel = new File("src/main/sample/SampleData-ExcelToJSON.xlsx");  // For testing purpose the excel file is kept under src/main/sample folder, in actual setup the files should be kep outside application folders.
        ExcelToJSONConverter converter = new ExcelToJSONConverter();
        JSONObject data = converter.excelToJson(excel);
        System.out.println("Excel file contains the Data:\n" + data);
    }

    public JSONObject excelToJson(File excel) {

        // hold the excel data sheet wise
        FileInputStream excelFile = null;
        Workbook workbook = null;

        try {
            // Creating file input stream
            excelFile = new FileInputStream(excel);

            String filename = excel.getName().toLowerCase();
            if (filename.endsWith(".xls") || filename.endsWith(".xlsx")) {
                // creating workbook object based on excel file format
                if (filename.endsWith(".xls")) {
                    workbook = new HSSFWorkbook(excelFile);
                } else {
                    workbook = new XSSFWorkbook(excelFile);
                }

                JSONObject jsonData = new JSONObject();
                JSONArray jsonArray = new JSONArray();

                // Reading each sheet one by one
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();

                    Iterator<Row> rowIterator = sheet.iterator();

                    // Assuming the header rows have two levels of sub-headers
                    Row headerRow1 = rowIterator.next();
                    Row headerRow2 = rowIterator.next();

                    while (rowIterator.hasNext()) {
                        Row dataRow = rowIterator.next();
                        JSONObject jsonObject = new JSONObject();

                        for (int j = 0; j < headerRow1.getLastCellNum(); j++) {
                            Cell headerCell1 = headerRow1.getCell(j); // Header Row 1
                            Cell headerCell2 = headerRow2.getCell(j); // Header Row 2
                            Cell dataCell = dataRow.getCell(j);

                            String header1 = headerCell1.getStringCellValue();
                            String header2 = headerCell2.getStringCellValue();
                            // String value = dataCell.getStringCellValue(); // You can handle other cell types as needed
                            String value = "";
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(dataCell.getNumericCellValue());
                            }

                            if (!header1.isEmpty()) {
                                JSONObject header1Object = jsonObject.optJSONObject(header1);
                                if (header1Object == null) {
                                    header1Object = new JSONObject();
                                    jsonObject.put(header1, header1Object);
                                }
                                header1Object.put(header2, value);
                            } else {
                                jsonObject.put(header2, value);
                            }
                        }

                        jsonArray.put(jsonObject);
                    }

                    jsonData.put(sheetName, jsonArray);
                }

                System.out.println(jsonData.toString(2)); // Print formatted JSON to console

                workbook.close();
                excelFile.close();

                return jsonData;
            } else {
                throw new IllegalArgumentException("File format not supported.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (excelFile != null) {
                try {
                    excelFile.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

}
