//package com.minitest2;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.nio.file.Paths;
//import java.text.SimpleDateFormat;
//import java.util.Date;
//import java.util.Scanner;
//
//import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ExcelWriter {
//
//    private Workbook workbook;
//    private Sheet sheet;
//    private File excelFile;
//
//    public ExcelWriter(File excelFile) throws InvalidFormatException, IOException {
//        this.excelFile = excelFile;
//        FileInputStream fis = new FileInputStream(excelFile);
//        workbook = new XSSFWorkbook(fis);
//        sheet = workbook.getSheetAt(0); // Assuming data is written to the first sheet
//        fis.close();
//    }
//
//    public void addMatch(String weightClass, String round, String athleteBlue, String athleteRed) throws IOException {
//        int lastRowIndex = sheet.getLastRowNum();
//
//        Row newRow = sheet.createRow(lastRowIndex + 1);
//
//        Cell cell = newRow.createCell(0, CellType.STRING);
//        cell.setCellValue(generateMatchCode());
//
//        cell = newRow.createCell(1, CellType.STRING);
//        cell.setCellValue(weightClass);
//
//        cell = newRow.createCell(2, CellType.STRING);
//        cell.setCellValue(round);
//
//        cell = newRow.createCell(3, CellType.STRING);
//        cell.setCellValue(athleteBlue);
//
//        cell = newRow.createCell(5, CellType.STRING);
//        cell.setCellValue(athleteRed);
//
//        FileOutputStream fos = new FileOutputStream(excelFile);
//        workbook.write(fos);
//        fos.close();
//
//        System.out.println("Match added successfully.");
//    }
//
//    private String generateMatchCode() {
//        // Logic to generate match code based on current maximum match code in the sheet
//        int maxMatchCode = 0;
//
//        for (Row row : sheet) {
//            if (row.getRowNum() == 0) continue; // Skip header row
//            Cell cell = row.getCell(0);
//            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
//                int currentCode = (int) cell.getNumericCellValue();
//                if (currentCode > maxMatchCode) {
//                    maxMatchCode = currentCode;
//                }
//            }
//        }
//
//        return String.format("%d", maxMatchCode + 1);
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
//        ExcelWriter writer = new ExcelWriter(excelFile);
//
//        System.out.println("Enter details for the new match:");
//        System.out.print("Weight Class: ");
//        String weightClass = scanner.nextLine();
//
//        System.out.print("Round: ");
//        String round = scanner.nextLine();
//
//        System.out.print("Athlete (Blue): ");
//        String athleteBlue = scanner.nextLine();
//
//        System.out.print("Athlete (Red): ");
//        String athleteRed = scanner.nextLine();
//
//        writer.addMatch(weightClass, round, athleteBlue, athleteRed);
//
//        scanner.close();
//    }
//}
