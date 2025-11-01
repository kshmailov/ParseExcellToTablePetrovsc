package org.example.utils;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.dto.StringTable;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.CellType.*;

@Slf4j
public class ParseExcellTable {
    @Getter
    @Setter
    private int idString;
    private final List<StringTable> tableStrings;

    public ParseExcellTable() {
        this.tableStrings = new ArrayList<>();
        this.idString = 0;
    }

    public ParseExcellTable(int idString) {
        this.tableStrings = new ArrayList<>();
        this.idString = idString;
    }

    /**
     * Основной метод чтения Excel-файла.
     */
    public List<StringTable> parseExcell(String fileName, boolean isNotConsistCspaString) {
        Path path = Paths.get(fileName);
        log.info("Открываю рабочую книгу {}", fileName);

        try (FileInputStream fis = new FileInputStream(path.toFile());
             Workbook workbook = new XSSFWorkbook(fis)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            log.info("Рабочая книга содержит {} лист(ов)", numberOfSheets);

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                log.info("Обрабатываю лист: {}", sheet.getSheetName());
                int rowNum =0;
                while (rowNum<=sheet.getLastRowNum()) {
                    Row row = sheet.getRow(rowNum);
                    if (row == null){
                        rowNum++;
                        continue;
                    }

                    Cell firstCell = row.getCell(0);
                    if (firstCell != null && (firstCell.getCellType() == NUMERIC||firstCell.getCellType() == FORMULA)) {
                        int endMergedRow = getMergedRegionEndRow(sheet,rowNum);
                        rowNum=endMergedRow+1;
                        handleRowData(row, sheet,endMergedRow, isNotConsistCspaString);
                    }else rowNum++;

                }
            }

        } catch (IOException e) {
            log.error("Ошибка при открытии рабочей книги {}: {}", fileName, e.getMessage(), e);
        }

        return tableStrings;
    }

    /**
     * Обработка строки с данными таблицы.
     */
    private void handleRowData(Row row, Sheet sheet, int endUvRow, boolean isNotConsistCspa) {
        StringTable currentTable = new StringTable();
        currentTable.setId(this.idString++);
        currentTable.setShm(getCellValueAsString(row.getCell(2)));
        currentTable.setTs(normalizeValue(row.getCell(3)));
        currentTable.setFormula(getCellValueAsString(row.getCell(4)));
        currentTable.setPo(getCellValueAsString(row.getCell(5)));
        currentTable.setSlice(getCellValueAsString(row.getCell(7)));
        String sign = switch (getCellValueAsString(row.getCell(8))){
            case "+" -> "0";
            case "-" -> "1";
            default -> "";
        };
        currentTable.setSign(sign);
        currentTable.setKpr(normalizeValue(row.getCell(9)));

        // определяем границы объединения (по первой колонке)
        int startRow =row.getRowNum();
        int startUvRow = startRow + 1;


        int colStepNumber = 10; // номер ступени
        int colUvList = 11;     // список УВ

        if(startRow!=endUvRow) tableUv(sheet, startUvRow, endUvRow, colStepNumber, colUvList, currentTable);

        // добавляем итоговую строку
        tableStrings.add(currentTable);
        log.info("Добавлена строка: {}", currentTable);

        if (isNotConsistCspa) {
            createCspaString(currentTable);
        }
    }

    /**
     * Чтение блока УВ (Номер ступени + Список УВ)
     * — добавляет все найденные УВ в одну запись StringTable
     */
    private void tableUv(Sheet sheet, int startRow, int endRow, int colStepNumber, int colUvList, StringTable table) {
        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) continue;

            Cell stepCell = row.getCell(colStepNumber);
            Cell uvCell = row.getCell(colUvList);

            String stepVal = getCellValueAsString(stepCell).strip();
            String uvVal = getCellValueAsString(uvCell).strip();

            if (stepCell == null || stepVal.isEmpty() ||
                    uvCell == null || uvVal.isEmpty()|| "-".equals(stepVal)
                    || "-".equals(uvVal)) break;

            String[] uvString = uvVal.split(";");
            StringBuilder uvBuilder = new StringBuilder();
            uvBuilder.append("s").append(stepVal).append("=");
            for (int i=0; i<uvString.length;i++){
                if (i>0) uvBuilder.append(" ");
                String uv = uvString[i].strip();
                uvBuilder.append(uv);
            }
            table.addUv(uvBuilder.toString());
            log.debug("Добавлено УВ: {}", uvBuilder);
        }
    }

    /**
     * Определяет последнюю строку объединения для заданной ячейки.
     * Если ячейка не объединена — возвращает тот же индекс строки.
     */
    private int getMergedRegionEndRow(Sheet sheet, int rowIndex) {
        if (rowIndex==sheet.getLastRowNum()) return rowIndex;
        int endRow =rowIndex;
        for (int i = endRow+1; i < sheet.getLastRowNum(); i++) {
            CellType cellType = sheet.getRow(i).getCell(0).getCellType();
            if (cellType==NUMERIC||cellType==FORMULA)return i-1;
            endRow++;
        }
        return endRow; // не объединена
    }

    /**
     * Возвращает значение ячейки в виде строки.
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            case FORMULA -> {
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue value = evaluator.evaluate(cell);
                    yield value != null ? value.formatAsString() : "";
                } catch (Exception e) {
                    log.warn("Ошибка вычисления формулы в ячейке {}: {}", cell.getAddress(), e.getMessage());
                    yield "";
                }
            }
            default -> "";
        };
    }

    /**
     * Нормализация значения (замена [Не задан] на пустую строку).
     */
    private String normalizeValue(Cell cell) {
        String val = getCellValueAsString(cell);
        return "[Не задан]".equals(val) ? "" : val;
    }

    /**
     * Создание дополнительной строки для ЦСПА.
     */
    private void createCspaString(StringTable originalTable) {
        StringTable additionalTable = new StringTable();
        additionalTable.setId(this.idString++);
        additionalTable.setShm(originalTable.getShm());
        additionalTable.setTs(originalTable.getTs());
        additionalTable.setFormula(originalTable.getFormula());
        additionalTable.setPo(originalTable.getPo());
        additionalTable.setSlice(originalTable.getSlice());
        additionalTable.setSign(originalTable.getSign());
        additionalTable.setKpr("*ФСч");
        additionalTable.addUv("s1=*ОН_ЛАПНУ");
        tableStrings.add(additionalTable);
        log.info("Добавлена строка для ЦСПА: {}", additionalTable);
    }
}
