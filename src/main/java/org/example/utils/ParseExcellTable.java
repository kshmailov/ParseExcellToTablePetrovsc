package org.example.utils;

import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.dto.StringTable;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
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
    public ParseExcellTable(){
        this.tableStrings = new ArrayList<>();
        this.idString=0;
    }
    public ParseExcellTable(int idString){
        this.tableStrings = new ArrayList<>();
        this.idString=idString;
    }


    /*public List<StringTable> parseExcell(String fileName , boolean isNotConsistCspaString){
        Path path = Paths.get(fileName);
        log.info("Открываю рабочую книгу {}", fileName);
        try (InputStream is = Files.newInputStream(path);
            Workbook workbook = WorkbookFactory.create(is)) {
            int numberOfSheets = workbook.getNumberOfSheets();
            log.info("Рабочая книга содержит листов: {}", numberOfSheets);
            List<Sheet> sheets = new ArrayList<>();
            for (int i=0; i<numberOfSheets; i++){
                sheets.add(workbook.getSheetAt(i));
                log.info("В рабочую книгу добавлен лист: {}", workbook.getSheetAt(i).getSheetName());
            }
            for (Sheet sheet : sheets){
                boolean isStringTableCreate=false;
                boolean isStringTableAdd=false;
                StringTable stringTable1 = null;
                int numberTableString=0;
                for (Row row : sheet){
                    Cell cell = row.getCell(10);
                    CellType numberUv = cell.getCellType();
                    CellStyle style = cell.getCellStyle();
                    Font font = workbook.getFontAt(style.getFontIndex());
                    if (isStringTableCreate && stringTable1 != null && (numberUv.equals(NUMERIC) || numberUv.equals(CellType.FORMULA)) && !font.getStrikeout()) {

                        StringBuilder stringBuilder = new StringBuilder();
                        int st;
                        if (numberUv.equals(NUMERIC)) {
                            st = (int) cell.getNumericCellValue();
                        } else {
                            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                            CellValue evaluatedValue = evaluator.evaluate(cell);
                            st = (int) evaluatedValue.getNumberValue();
                        }
                        stringBuilder.append("s").append(st).append("=");
                        String uv = row.getCell(11).getStringCellValue();
                        String[] uvList = uv.split("; ");
                        for (int i = 0; i < uvList.length; i++) {
                            if (i == 0) {
                                stringBuilder.append(uvList[i]);
                                continue;
                            }
                            stringBuilder.append(" ").append(uvList[i]);
                        }
                        stringTable1.addUv(stringBuilder.toString());
                        continue;

                    }
                    if (stringTable1!=null&&!isStringTableAdd){
                        tableStrings.add(stringTable1);
                        isStringTableAdd=true;
                        log.info("Добавлена строка {}", stringTable1);
                        isStringTableCreate=false;
                        setIdString(getIdString()+1);
                        if (isNotConsistCspaString) createCspaString(stringTable1);

                    }
                    Cell cellNumberString = row.getCell(0);
                    CellType typeNumberString = cellNumberString.getCellType();
                    if (typeNumberString.equals(NUMERIC)||typeNumberString.equals(FORMULA)){
                        if (typeNumberString.equals(NUMERIC)){
                            numberTableString=(int) cellNumberString.getNumericCellValue();
                        }else {
                            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                            CellValue evaluatedValue = evaluator.evaluate(cellNumberString);
                            numberTableString = (int) evaluatedValue.getNumberValue();
                        }
                        if (stringTable1!=null&&numberTableString!=getIdString()+1&&stringTable1.getUvList().isEmpty()){
                            tableStrings.add(stringTable1);
                            log.info("Добавлена строка без УВ {}", stringTable1);
                            setIdString(getIdString()+1);
                            if (isNotConsistCspaString) createCspaString(stringTable1);
                        }
                        isStringTableCreate=true;
                        stringTable1 = new StringTable();
                        int id = getIdString();
                        stringTable1.setId(id);
                        String scheme = row.getCell(2).getStringCellValue();
                        stringTable1.setShm(scheme);
                        String ts = row.getCell(3).getStringCellValue();
                        if (!ts.equals("[Не задан]")) stringTable1.setTs(ts);
                        String formula = row.getCell(4).getStringCellValue();
                        stringTable1.setFormula(formula);
                        String po = row.getCell(5).getStringCellValue();
                        stringTable1.setPo(po);
                        String slc = row.getCell(7).getStringCellValue();
                        stringTable1.setSlice(slc);
                        String sign = row.getCell(8).getStringCellValue();
                        stringTable1.setSign(sign);
                        String kpr = row.getCell(9).getStringCellValue();
                        if (!kpr.equals("[Не задан]")) stringTable1.setKpr(kpr);
                        isStringTableAdd=false;
                        if (row.getRowNum()==sheet.getLastRowNum()){
                            tableStrings.add(stringTable1);
                            log.info("Добавлена строка без УВ {}", stringTable1);
                            setIdString(getIdString()+1);
                            if (isNotConsistCspaString) createCspaString(stringTable1);
                        }
                        continue;
                    }
                }
            }


        } catch (IOException e) {
            log.error("Ошибка при открытии рабочей книги {}: {}", fileName, e.getMessage(), e);
        }

        return tableStrings;
    }
    private void createCspaString(StringTable stringTable){
        StringTable stringTable2 = new StringTable();
        stringTable2.setId(getIdString());
        stringTable2.setShm(stringTable.getShm());
        stringTable2.setTs(stringTable.getTs());
        stringTable2.setFormula(stringTable.getFormula());
        stringTable2.setPo(stringTable.getPo());
        stringTable2.setSlice(stringTable.getSlice());
        stringTable2.setSign(stringTable.getSign());
        stringTable2.setKpr("*ФСч");
        String uvSttringTable2 = "s1=*ОН_ЛАПНУ";
        stringTable2.addUv(uvSttringTable2);
        tableStrings.add(stringTable2);
        log.info("Добавлена строка для возврата признака работы из ЛАПНУ в ЦСПА {}", stringTable2);
        setIdString(getIdString()+1);
    }*/

    /**
     * Парсим таблицу с применением streaming API для эффективного потребления памяти.
     */
    public List<StringTable> parseExcell(String fileName, boolean isNotConsistCspaString) {
        Path path = Paths.get(fileName);
        log.info("Открываю рабочую книгу {}", fileName);

        try (
             SXSSFWorkbook workbook = new SXSSFWorkbook(new FileInputStream(fileName))) { // Используем streaming для снижения нагрузки на память

            int numberOfSheets = workbook.getNumberOfSheets();
            log.info("Рабочая книга содержит листов: {}", numberOfSheets);

            for (int i = 0; i < numberOfSheets; i++) {
                SXSSFSheet sheet = workbook.getSheetAt(i);
                log.info("Обрабатываю лист: {}", sheet.getSheetName());

                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    SXSSFRow row = sheet.getRow(rowNum);

                    if (row == null) {
                        continue; // Пропускаем пустые строки
                    }

                    // Проверяем наличие значимой информации в строке
                    Cell firstCell = row.getCell(0);
                    if (firstCell != null && firstCell.getCellType() == NUMERIC) {
                        handleRowData(row, sheet, isNotConsistCspaString);
                    }
                }
            }
//        } catch (IOException  e) {
//            log.error("Ошибка при открытии рабочей книги {}: {}", fileName, e.getMessage(), e);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }

        return tableStrings;
    }

    /**
     * Метод обрабатывает одну строку таблицы и добавляет соответствующие записи в список таблиц.
     *
     * @param row              Текущая строка
     * @param sheet            Лист, содержащий строку
     * @param isNotConsistCspa Флаг для дополнительной обработки строк
     */
    private void handleRowData(SXSSFRow row, SXSSFSheet sheet, boolean isNotConsistCspa) {
        StringTable currentTable = new StringTable();
        currentTable.setId(this.idString++);

        // Заполняем данные таблицы
        currentTable.setShm(getCellValueAsString(row.getCell(2)));
        currentTable.setTs(getCellValueAsString(row.getCell(3)).equals("[Не задан]") ? "" : getCellValueAsString(row.getCell(3)));
        currentTable.setFormula(getCellValueAsString(row.getCell(4)));
        currentTable.setPo(getCellValueAsString(row.getCell(5)));
        currentTable.setSlice(getCellValueAsString(row.getCell(7)));
        currentTable.setSign(getCellValueAsString(row.getCell(8)));
        currentTable.setKpr(getCellValueAsString(row.getCell(9)).equals("[Не задан]") ? "" : getCellValueAsString(row.getCell(9)));

        // Проверяем специфичные условия для добавления записей UV
        Cell uvCell = row.getCell(10);
        if (uvCell != null && (uvCell.getCellType() == NUMERIC || uvCell.getCellType() == FORMULA)) {
            addUVToTable(currentTable, uvCell);
        }

        // Добавляем запись в общий список
        tableStrings.add(currentTable);

        // Дополнительная обработка строк для случая отсутствия CSPA
        if (isNotConsistCspa) {
            createCspaString(currentTable);
        }
    }

    /**
     * Возвращает значение ячейки как строку.
     *
     * @param cell Ячейка, значение которой нужно преобразовать
     * @return Строковое представление значения ячейки
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().toString();
            case NUMERIC:
                return Double.toString(cell.getNumericCellValue());
            default:
                return "";
        }
    }

    /**
     * Добавляет UV-значение в объект StringTable.
     *
     * @param table Таблица, куда добавляется UV
     * @param uvCell Ячейка, содержащая UV-значение
     */
    private void addUVToTable(StringTable table, Cell uvCell) {
        StringBuilder builder = new StringBuilder();
        double numericValue = uvCell.getNumericCellValue();
        int uvNumber = (int) Math.round(numericValue);
        builder.append("s").append(uvNumber).append("=");

        Cell uvTextCell = uvCell.getRow().getCell(11);
        if (uvTextCell != null && uvTextCell.getCellType() == STRING) {
            String[] uvValues = uvTextCell.getStringCellValue().split(";");
            for (int i = 0; i < uvValues.length; i++) {
                if (i > 0) {
                    builder.append(" ");
                }
                builder.append(uvValues[i].trim());
            }
        }
        table.addUv(builder.toString());
    }

    /**
     * Создание дополнительной строки для учета CPSA-кодов.
     *
     * @param originalTable Исходная таблица
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
        log.info("Добавлена дополнительная строка для учёта CPSA-кода: {}", additionalTable);
    }

}
