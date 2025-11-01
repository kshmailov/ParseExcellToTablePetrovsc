package org.example.utils;

import lombok.extern.slf4j.Slf4j;
import org.example.dto.StringTable;

import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
@Slf4j
public class CreatePau {
    private List<StringTable> tableList;
    public CreatePau (List<StringTable> stringTables){
        tableList = new ArrayList<>();
        tableList.addAll(stringTables);
    }
    public void writeToPau(String path){
        Charset windows1251 = Charset.forName("Windows-1251");
        try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(path),windows1251))){
            StringBuilder builder = new StringBuilder();
            int i=1;
            for (StringTable currentTable : tableList){
                if (i==100) {
                    log.info("Записываем секции\n{}",builder);
                    writer.write(builder.toString());
                    builder.setLength(0);
                }
                builder.append("[ТУВ_").append(currentTable.getId()).append("]").append("\r\n");
                builder.append("\t").append("class=СтрокаУТ").append("\r\n");
                builder.append("\t").append("shm=").append(currentTable.getShm()).append("\r\n");
                String ts = currentTable.getTs();
                if (!ts.isEmpty()) builder.append("\t").append("ts=").append(ts).append("\r\n");
                String formula = currentTable.getFormula();
                if (!formula.isEmpty()) builder.append("\t").append("formula=").append(formula).append("\r\n");
                builder.append("\t").append("po=").append(currentTable.getPo()).append("\r\n");
                builder.append("\t").append("slice=").append(currentTable.getSlice()).append("\r\n");
                String sign = currentTable.getSign();
                if(!sign.isEmpty()) builder.append("\t").append("sign=").append(sign).append("\r\n");
                String kpr = currentTable.getKpr();
                if (!kpr.isEmpty()) builder.append("\t").append("kpr=").append(kpr).append("\r\n");
                List<String> uvList = new ArrayList<>(currentTable.getUvList());
                if (!uvList.isEmpty()) for (String uv : uvList) builder.append("\t").append(uv).append("\r\n");
                builder.append("\r\n");
                i++;
            }
            if(!builder.isEmpty()){
                log.info("Записываем секции\n{}",builder);
                writer.write(builder.toString());
                builder.setLength(0);
            }
            log.info("Запись закончена!");
        } catch (IOException e) {
            log.error("Ошибка записи файла {}: {}", path, e.getMessage(), e);
        }
    }
}
