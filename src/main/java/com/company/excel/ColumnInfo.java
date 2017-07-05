package com.company.excel;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ColumnInfo {
    public String value;
    public String typeName;
    public int columnNumber;

    public void printResults() {
        String line;
        line = "%s %s %s";
        line = String.format(line,this.value,this.typeName,this.columnNumber);
        System.out.println(line);
    }
}
