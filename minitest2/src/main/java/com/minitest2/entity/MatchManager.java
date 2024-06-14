package com.minitest2.entity;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Scanner;

public class MatchManager {
    private static ExcelReader reader;
    private static ExcelWriter writer;
    private static File excelFile;

    public static void main(String[] args) throws InvalidFormatException, IOException {
        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("Please enter the name of the Excel file (e.g., Matches.xlsx):");
            String fileName = scanner.nextLine();

            // Assuming the file is in the Downloads directory
            String userHome = System.getProperty("user.home");
            excelFile = Paths.get(userHome, "Downloads", fileName).toFile();

            if (excelFile.exists()) {
                break;
            } else {
                System.out.println("The specified file does not exist in the Downloads directory. Please try again.");
            }
        }

        reader = new ExcelReader(excelFile);
        writer = new ExcelWriter(excelFile);

        int choice;
        do {
            System.out.println("\nMenu:");
            System.out.println("1. Load Matches from Excel");
            System.out.println("2. Save Matches to Excel");
            System.out.println("3. Add Match");
            System.out.println("4. Update Match");
            System.out.println("5. Search Matches by Weight Class");
            System.out.println("6. Display All Matches");
            System.out.println("7. Exit");
            System.out.print("Enter your choice: ");
            choice = scanner.nextInt();
            scanner.nextLine(); // Consume newline character

            switch (choice) {
                case 1:
                    reader.readFromExcelFile();
                    break;
                case 2:
                    saveMatchesToExcel(scanner);
                    break;
                case 3:
                    addMatch(scanner);
                    break;
                case 4:
                    updateMatch(scanner);
                    break;
                case 5:
                    searchMatchesByWeightClass(scanner);
                    break;
                case 6:
                    reader.readFromExcelFile();
                    break;
                case 7:
                    System.out.println("Exiting...");
                    break;
                default:
                    System.out.println("Invalid choice. Please enter a number from 1 to 7.");
            }
        } while (choice != 7);

        scanner.close();
        reader.close(); // Close workbook
    }

    private static void saveMatchesToExcel(Scanner scanner) throws IOException {
        System.out.println("Enter the name of the Excel file to save (e.g., Matches_new.xlsx):");
        String fileName = scanner.nextLine();

        File outputFile = Paths.get(System.getProperty("user.home"), "Downloads", fileName).toFile();

        // Copy content from current Excel file to the new file
        try {
            reader = new ExcelReader(excelFile);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
        ExcelWriter.saveMatchesToExcel(outputFile, reader.getSheet());

        System.out.println("Matches saved to " + outputFile.getAbsolutePath());
    }

    private static void addMatch(Scanner scanner) throws IOException {
        System.out.print("Weight Class: ");
        String weightClass = scanner.nextLine();

        System.out.print("Round: ");
        String round = scanner.nextLine();

        System.out.print("Athlete (Blue) - Name: ");
        String nameBlue = scanner.nextLine();
        System.out.print("Athlete (Blue) - Unit: ");
        String unitBlue = scanner.nextLine();

        System.out.print("Athlete (Red) - Name: ");
        String nameRed = scanner.nextLine();
        System.out.print("Athlete (Red) - Unit: ");
        String unitRed = scanner.nextLine();

        writer.addMatch(weightClass, round, nameBlue, unitBlue, nameRed, unitRed);
    }

    private static void updateMatch(Scanner scanner) throws IOException {
        System.out.print("Enter match code to update: ");
        int matchCode = scanner.nextInt();
        scanner.nextLine(); // Consume newline character

        boolean matchFound = false;

        for (Row row : reader.getSheet()) {
            Cell cell = row.getCell(0); // Assuming match code is in the first column

            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                int currentMatchCode = (int) cell.getNumericCellValue();

                if (currentMatchCode == matchCode) {
                    // Match found, update details
                    System.out.print("New Weight Class: ");
                    String weightClass = scanner.nextLine();

                    System.out.print("New Round: ");
                    String round = scanner.nextLine();

                    System.out.print("New Athlete (Blue) - Name: ");
                    String nameBlue = scanner.nextLine();
                    System.out.print("New Athlete (Blue) - Unit: ");
                    String unitBlue = scanner.nextLine();

                    System.out.print("New Athlete (Red) - Name: ");
                    String nameRed = scanner.nextLine();
                    System.out.print("New Athlete (Red) - Unit: ");
                    String unitRed = scanner.nextLine();

                    // Update match details in Excel
                    writer.updateMatch(row.getRowNum(), weightClass, round,
                            nameBlue, unitBlue, nameRed, unitRed);

                    System.out.println("Match updated successfully.");
                    matchFound = true;
                    break;
                }
            }
        }

        if (!matchFound) {
            System.out.println("Match with code " + matchCode + " not found.");
        }
    }

    private static void searchMatchesByWeightClass(Scanner scanner) {
        System.out.print("Enter weight class to search: ");
        String weightClass = scanner.nextLine();

        boolean matchesFound = false;

        for (Row row : reader.getSheet()) {
            Cell cell = row.getCell(1); // Assuming weight class is in the second column

            if (cell != null) {
                if (cell.getCellType() == CellType.STRING) {
                    String currentWeightClass = cell.getStringCellValue();

                    if (currentWeightClass.equalsIgnoreCase(weightClass)) {
                        // Match found, display match details
                        Cell cellMatchCode = row.getCell(0);
                        Cell cellRound = row.getCell(2);
                        Cell cellAthleteBlue = row.getCell(3);
                        Cell cellAthleteRed = row.getCell(4);

                        System.out.printf("Match Code: %d%n", (int) cellMatchCode.getNumericCellValue());
                        System.out.printf("Round: %s%n", cellRound.getStringCellValue());
                        System.out.printf("Athlete (Blue): %s%n", cellAthleteBlue.getStringCellValue());
                        System.out.printf("Athlete (Red): %s%n", cellAthleteRed.getStringCellValue());

                        matchesFound = true;
                    }
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    // Handle numeric cell if needed
                    // Example: int currentWeightClass = (int) cell.getNumericCellValue();
                }
            }
        }

        if (!matchesFound) {
            System.out.println("No matches found for weight class " + weightClass);
        }
    }

}
