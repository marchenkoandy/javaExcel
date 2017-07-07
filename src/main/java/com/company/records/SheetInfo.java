package com.company.records;

import com.company.Main;

import java.util.ArrayList;

/**
 * Created by user on 06.07.2017.
 */
public class SheetInfo implements ExcelInterface {
    public String                       name;
    public ArrayList<ColumnInfo>        columnInfos;
    public static int LEVEL             = Main.LEVEL + 2;

    public void                         info(){
        Main.printOnLevel(LEVEL,Main.delimiter());
        Main.printOnLevel(LEVEL,"Worksheet name:             " + name);
        Main.printOnLevel(LEVEL,"Worksheet columns count:    " + columnInfos.size());
        Main.printOnLevel(LEVEL,Main.delimiter());
    }
    public void                         printResults() {
        info();
        for (ColumnInfo info: columnInfos) {
            info.printResults();
        }
    }
}
