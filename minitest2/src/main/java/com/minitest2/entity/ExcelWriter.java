package com.minitest2.entity;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class ExcelWriter {
    private Workbook workbook;
    private Sheet sheet;
    private File excelFile;

    public ExcelWriter(File excelFile) throws InvalidFormatException, IOException {
        this.excelFile = excelFile;
        FileInputStream fis = new FileInputStream(excelFile);
        workbook = WorkbookFactory.create(fis);
        sheet = workbook.getSheetAt(0); // Assuming data is on the first sheet
        fis.close();
    }

    public static void saveMatchesToExcel(File outputFile, Sheet sheet) throws IOException {
        FileOutputStream fos = new FileOutputStream(outputFile);
        sheet.getWorkbook().write(fos);
        fos.close();
    }

    public void addMatch(String weightClass, String round,
                         String athleteBlueName, String athleteBlueUnit, String athleteRedName, String athleteRedUnit) throws IOException {
        int lastRowIndex = sheet.getLastRowNum();

        Row newRow = sheet.createRow(lastRowIndex + 1);

        Cell cell = newRow.createCell(0, CellType.NUMERIC);
        cell.setCellValue(generateMatchCode());

        cell = newRow.createCell(1, CellType.STRING);
        cell.setCellValue(weightClass);

        cell = newRow.createCell(2, CellType.STRING);
        cell.setCellValue(round);

        cell = newRow.createCell(3, CellType.STRING);
        cell.setCellValue(athleteBlueName);

        cell = newRow.createCell(4, CellType.STRING);
        cell.setCellValue(athleteBlueUnit);

        cell = newRow.createCell(5, CellType.STRING);
        cell.setCellValue(athleteRedName);

        cell = newRow.createCell(6, CellType.STRING);
        cell.setCellValue(athleteRedUnit);
        FileOutputStream fos = new FileOutputStream(excelFile);
        workbook.write(fos);
        fos.close();

        System.out.println("Match added successfully.");
    }

    private int generateMatchCode() {
        int maxMatchCode = 0;

        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            if (currentRow.getRowNum() == 0) continue; // Skip header row
            Cell cell = currentRow.getCell(0);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                int currentCode = (int) cell.getNumericCellValue();
                if (currentCode > maxMatchCode) {
                    maxMatchCode = currentCode;
                }
            }
        }

        return maxMatchCode + 1;
    }

    public void updateMatch(int rowIndex, String weightClass, String round,
                            String athleteBlueName, String athleteBlueUnit,
                            String athleteRedName, String athleteRedUnit) throws IOException {
        Row row = sheet.getRow(rowIndex);

        Cell cell = row.createCell(1, CellType.STRING);
        cell.setCellValue(weightClass);

        cell = row.createCell(2, CellType.STRING);
        cell.setCellValue(round);

        cell = row.createCell(3, CellType.STRING);
        cell.setCellValue(athleteBlueName + " - " + athleteBlueUnit);

        cell = row.createCell(4, CellType.STRING);
        cell.setCellValue(athleteRedName + " - " + athleteRedUnit);

        FileOutputStream fos = new FileOutputStream(excelFile);
        workbook.write(fos);
        fos.close();

        System.out.println("Match updated successfully.");
    }

    public Sheet getSheet() {
        return sheet;
    }
}
