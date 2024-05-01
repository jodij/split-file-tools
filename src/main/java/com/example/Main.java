package com.example;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Slf4j
public class Main {

    private static final String FILE_DIR = "./data";
    private static final Integer SPLIT_FILE_LIMIT = 500;

    public static void main(String[] args) {
        System.out.println("Process Started!");
        String fileName = "STOK RMA.xlsx";

        readExcel(FILE_DIR + File.separator + fileName);
        System.out.println("Process Done!");
    }

    public static void readExcel(String filePath) {
        File file = new File(filePath);
        try {
            FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(inputStream);

            for (Sheet sheet : workbook) {
                int firstRow = sheet.getFirstRowNum();
                int lastRow = sheet.getLastRowNum();

                System.out.println("Total Row: " + lastRow);

                int fileCount = 1;
                List<Row> rows = new ArrayList<>();

                Row headerRow = sheet.getRow(0);

                for (int index = firstRow + 1; index <= lastRow; index++) {
                    Row row = sheet.getRow(index);
                    rows.add(row);
                    if (index % SPLIT_FILE_LIMIT == 0) {
                        writeExcel(FILE_DIR + File.separator + "output_" + fileCount + ".xlsx", rows, headerRow);
                        System.out.println("file ke: " + fileCount + ", row count: " + rows.size());
                        fileCount++;
                        rows.clear();
                    }
                }
                /* last file */
                writeExcel(FILE_DIR + File.separator + "output_" + fileCount + ".xlsx", rows, headerRow);
                System.out.println("file ke: " + fileCount + ", row count: " + rows.size());
                rows.clear();
            }

            inputStream.close();
        } catch (IOException e) {
            log.error("ERROR readExcel, {}", e.getMessage());
        }
    }

    public static void writeExcel(String fileLocation, List<Row> rows, Row headerRow) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("Data");

            Row header = sheet.createRow(0);
            XSSFFont font = workbook.createFont();
            font.setFontName("Arial");
            font.setFontHeightInPoints((short) 16);
            font.setBold(true);

            Iterator<Cell> cellHeaderIterator = headerRow.cellIterator();
            while (cellHeaderIterator.hasNext()) {
                Cell cell = cellHeaderIterator.next();
                Cell headerCell = header.createCell(cell.getColumnIndex());
                headerCell.setCellValue(headerRow.getCell(cell.getColumnIndex()).toString());
            }

            int rowNumber = 1;

            for (Row row : rows) {
                Row newRow = sheet.createRow(rowNumber);
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    Cell newCell = newRow.createCell(cell.getColumnIndex());
                    switch (cell.getCellType()) {
                        case STRING:
                            newCell.setCellValue(row.getCell(cell.getColumnIndex()).getStringCellValue());
                            break;
                        case NUMERIC:
                            newCell.setCellValue(row.getCell(cell.getColumnIndex()).toString());
                            break;
                        case BOOLEAN:
                            newCell.setCellValue(row.getCell(cell.getColumnIndex()).getBooleanCellValue());
                            break;
                        case FORMULA:
                            newCell.setCellValue(row.getCell(cell.getColumnIndex()).getCellFormula());
                            break;
                        default:
                            newCell.setCellValue(row.getCell(cell.getColumnIndex()).toString());
                    }

                }
                rowNumber++;
            }

            /* auto size column */
            while (cellHeaderIterator.hasNext()) {
                Cell cell = cellHeaderIterator.next();
                sheet.autoSizeColumn(cell.getColumnIndex());

            }

            FileOutputStream outputStream = new FileOutputStream(fileLocation);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            log.error("ERROR writeExcel, {}", e.getMessage());
        }
    }
}

