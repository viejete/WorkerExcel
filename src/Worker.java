import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class Worker {

    public static void main(String[] args) {

        String fileName = "C:/Excel.xls";
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("FirstSheet");

        HSSFRow rowHead = sheet.createRow(0);
        rowHead.createCell(0).setCellValue("No.");
        rowHead.createCell(1).setCellValue("Name");
        rowHead.createCell(2).setCellValue("Address");
        rowHead.createCell(3).setCellValue("Email");

        HSSFRow row = sheet.createRow(1);
        row.createCell(0).setCellValue(1);
        row.createCell(1).setCellValue("Carlos");
        row.createCell(2).setCellValue("Spain");
        row.createCell(3).setCellValue("sistemas@egregia.net");

        try {
            FileOutputStream fileOut = new FileOutputStream(fileName);
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Your excel file has been generated!");
        } catch (IOException e) {
            e.printStackTrace();
        }

        secondExcel();

    }

    private static void secondExcel() {

        try {

            double sumaUPB = 0;
            double sumaUPBSolvency = 0;
            double sumaUPBInsolvency = 0;
            double tmp = 0;

            File excel = new File("C:\\Users\\Administrador\\Desktop\\test.xlsx");
            FileInputStream fis = new FileInputStream(excel);

            // Finds the workbook instance for XLSX file
            XSSFWorkbook wb = new XSSFWorkbook(fis);

            // Return first sheet from the XLSX workbook
            XSSFSheet firstSheet = wb.getSheetAt(0);

            // Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = firstSheet.iterator();

            // Travesing over each row of XLSX file
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    if (cell.getColumnIndex() == 23) {

                        try {
                            tmp = cell.getNumericCellValue();
                            sumaUPB += tmp;
                        } catch (IllegalStateException e) {
                            System.err.println("String");
                        }

                    } else if (cell.getColumnIndex() == 25) {

                        if (cell.getStringCellValue().equalsIgnoreCase("yes")) {
                            sumaUPBInsolvency += tmp;
                        } else if (cell.getStringCellValue().equalsIgnoreCase("no")) {
                            sumaUPBSolvency += tmp;
                        }
                    }
                }
            }

            System.out.println("Total amount -> " + sumaUPB);
            System.out.println("Total amount without insolvency -> " + sumaUPBSolvency);
            System.out.println("Total amount with insolvency -> " + sumaUPBInsolvency);
            System.out.println("Total amount sum -> " + (sumaUPBInsolvency + sumaUPBSolvency));


        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
