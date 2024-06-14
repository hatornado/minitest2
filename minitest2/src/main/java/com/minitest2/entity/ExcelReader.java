package com.minitest2.entity;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class ExcelReader {
    private Workbook workbook;
    private Sheet sheet;

    public ExcelReader(File excelFile) throws InvalidFormatException, IOException {
        this.workbook = WorkbookFactory.create(excelFile);
        this.sheet = workbook.getSheetAt(0); // Assuming data is on the first sheet
    }

    public void readFromExcelFile() {
        for (Row row : sheet) {
            for (Cell cell : row) {
                printCellValue(cell);
                System.out.print("\t");
            }
            System.out.println();
        }
    }

    private void printCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                System.out.print(cell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(new SimpleDateFormat("MM-dd-yyyy").format(cell.getDateCellValue()));
                } else {
                    System.out.print((int) cell.getNumericCellValue());
                }
                break;
            default:
                break;
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void close() throws IOException {
        workbook.close();
    }
}
