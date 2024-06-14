//package com.minitest2;
//
//import java.io.File;
//import java.io.IOException;
//import java.nio.file.Paths;
//import java.text.SimpleDateFormat;
//import java.util.Scanner;
//
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.DateUtil;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ExcelReader {
//
//    private Workbook workbook;
//    private Sheet sheet;
//
//    private ExcelReader(File excelFile) throws InvalidFormatException, IOException {
//        workbook = new XSSFWorkbook(excelFile);
//        sheet = workbook.getSheetAt(0);
//    }
//
//    public static void main(String[] args) throws InvalidFormatException, IOException {
//        Scanner scanner = new Scanner(System.in);
//        File excelFile = null;
//
//        while (true) {
//            System.out.println("Please enter the name of the Excel file (e.g., Matches.xlsx):");
//            String fileName = scanner.nextLine();
//
//            // Assuming the file is in the Downloads directory
//            String userHome = System.getProperty("user.home");
//            excelFile = Paths.get(userHome, "Downloads", fileName).toFile();
//
//            if (excelFile.exists()) {
//                break;
//            } else {
//                System.out.println("The specified file does not exist in the Downloads directory. Please try again.");
//            }
//        }
//
//        scanner.close();
//        ExcelReader reader = new ExcelReader(excelFile);
//        reader.readFromExcelFile();
//    }
//
//    private void readFromExcelFile() throws IOException {
//        for (Row row : sheet) {
//            System.out.println();
//
//            for (Cell cell : row) {
//                printCellValue(cell);
//                System.out.print("\t");
//            }
//        }
//
//        workbook.close();
//    }
//
//    private void printCellValue(Cell cell) {
//        switch (cell.getCellType()) {
//            case STRING:
//                System.out.print(cell.getStringCellValue());
//                break;
//            case NUMERIC:
//                if (DateUtil.isCellDateFormatted(cell)) {
//                    System.out.print(new SimpleDateFormat("MM-dd-yyyy").format(cell.getDateCellValue()));
//                } else {
//                    System.out.print((int) cell.getNumericCellValue());
//                }
//                break;
//            default:
//                break;
//        }
//    }
//}
