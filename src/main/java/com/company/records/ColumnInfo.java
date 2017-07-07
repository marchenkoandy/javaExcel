package com.company.records;

import com.company.Main;
import org.apache.poi.ss.usermodel.CellType;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ColumnInfo implements ExcelInterface {
    public String                               value;
    public CellType                             type;
    public int                                  columnNumber;
    public static int LEVEL                     = Main.LEVEL + 3;

    public void                                 info() {
//        empty for this class
    }
    public void                                 printResults() {
        String line;
        line = "%s %s %s";
        line = String.format(line,this.value,this.type.toString(),this.columnNumber);
        Main.printOnLevel(LEVEL,line);
    }
}
