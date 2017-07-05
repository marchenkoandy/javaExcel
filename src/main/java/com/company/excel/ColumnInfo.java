package com.company.excel;

import com.company.Main;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ColumnInfo {
    public String value;
    public String typeName;
    public int columnNumber;
    public static int LEVEL = 4;

    public void printResults() {
        String line;
        line = "%s %s %s";
        line = String.format(line,this.value,this.typeName,this.columnNumber);
        Main.printOnLevel(LEVEL,line);
    }
}
