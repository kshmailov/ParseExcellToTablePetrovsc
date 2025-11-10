package org.example.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
@Slf4j
public class CompareTables {
    public static void main(String[] args) {
        String file1 = "data/compare/Table1.xlsx";
        String file2 = "data/compare/Table2.xlsx";

        compareExcelFiles(file1, file2);
    }

    public static void compareExcelFiles(String filePath1, String filePath2)  {
        try (FileInputStream fis1 = new FileInputStream(filePath1);
             FileInputStream fis2 = new FileInputStream(filePath2);

             Workbook wb1 = new XSSFWorkbook(fis1);
             Workbook wb2 = new XSSFWorkbook(fis2)) {

            Sheet sheet1 = wb1.getSheetAt(0);
            Sheet sheet2 = wb2.getSheetAt(0);

            int maxRows = sheet1.getLastRowNum();

            log.info("Сравнение Excel файлов:\n");
            int n = 0;
            for (int i = 0; i <= maxRows; i++) {
                Row row1 = sheet1.getRow(i);
                Row row2 = sheet2.getRow(i);

                int maxCols = row1.getLastCellNum();

                for (int j = 0; j < maxCols; j++) {
                    Cell cell1 = (row1 == null) ? null : row1.getCell(j);
                    Cell cell2 = (row2 == null) ? null : row2.getCell(j);

                    String val1 = getCellValue(cell1);
                    String val2 = getCellValue(cell2);

                    if (!val1.equals(val2)) {
                        log.info("Различие: строка {}, столбец {} → {} ≠ {}",
                                i + 1, j + 1, val1, val2);
                    }
                }
            }

        } catch (IOException e) {
            log.error("Ошибка открытия файла {} или {}", filePath1, filePath2);
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
}
