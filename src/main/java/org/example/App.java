package org.example;

import org.example.dto.StringTable;
import org.example.utils.ParseExcellTable;

import java.util.ArrayList;
import java.util.List;

public class App
{
    public static void main( String[] args ) {
//        String input1 = "data/1.xlsx";
        String input2 = "data/2.xlsx";
        List<StringTable> stringTables = new ArrayList<>();
//        ParseExcellTable parseExcellTable1 = new ParseExcellTable();
//        stringTables.addAll(parseExcellTable1.parseExcell(input1,false));
//        int id = stringTables.get(stringTables.size()-1).getId();
//        parseExcellTable1 = new ParseExcellTable(id);
        ParseExcellTable parseExcellTable1 = new ParseExcellTable(13359);
        stringTables.addAll(parseExcellTable1.parseExcell(input2,true));

    }
}
