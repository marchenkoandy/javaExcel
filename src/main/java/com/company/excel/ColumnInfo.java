package com.company.excel;

import com.company.Main;
import org.apache.poi.ss.usermodel.CellType;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ColumnInfo {
    public String value;
//    public String typeName;
    public CellType type;
    public int columnNumber;
    public static int LEVEL = ExcelWorksheet.LEVEL + 1;

    public void printResults() {
        String line;
        line = "%s %s %s";
        line = String.format(line,this.value,this.type.toString(),this.columnNumber);
        Main.printOnLevel(LEVEL,line);
    }
}
