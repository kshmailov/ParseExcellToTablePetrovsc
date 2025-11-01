package org.example;

import org.example.dto.StringTable;
import org.example.utils.CreatePau;
import org.example.utils.ParseExcellTable;

import java.util.ArrayList;
import java.util.List;

public class App
{
    public static void main( String[] args ) {
        String input1 = "data/in/1.xlsx";
        String input2 = "data/in/2.xlsx";
        String out = "data/out/pau.ini";
        ParseExcellTable parseExcellTable1 = new ParseExcellTable();
        List<StringTable> stringTables = new ArrayList<>(parseExcellTable1.parseExcell(input1, false));
        int id = stringTables.getLast().getId()+1;
        parseExcellTable1 = new ParseExcellTable(id);
        stringTables.addAll(parseExcellTable1.parseExcell(input2,true));
        CreatePau createPau = new CreatePau(stringTables);
        createPau.writeToPau(out);

    }
}
