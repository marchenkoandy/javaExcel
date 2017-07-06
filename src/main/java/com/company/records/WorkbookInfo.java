package com.company.records;

import com.company.Main;

import java.util.ArrayList;

/**
 * Created by user on 06.07.2017.
 */
public class WorkbookInfo {
    public String                           name;
    public ArrayList<SheetInfo>             sheetInfos;
    public static int LEVEL                 = Main.LEVEL + 1;
    public void                             info(){
        Main.printOnLevel(LEVEL,Main.delimiter());
        Main.printOnLevel(LEVEL,"Workbook file name:     " + name);
        Main.printOnLevel(LEVEL,"Workbook sheets count:  " + sheetInfos.size());
        Main.printOnLevel(LEVEL,Main.delimiter());
    }
    public void                             printResults() {
        info();
        for (SheetInfo sheet:sheetInfos) {
            sheet.printResults();
        }
    }
}
