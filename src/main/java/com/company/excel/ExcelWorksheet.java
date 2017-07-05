package com.company.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ExcelWorksheet {
    private Row headerRow;
    private Row firstRow;
    private int headerRowColumnsCount;
    private Sheet sheet;
    private ArrayList<ColumnInfo> results = new ArrayList<ColumnInfo>();
    private void initializeSheet() {
        headerRow = sheet.getRow(0);
        firstRow = sheet.getRow(1);
        if (headerRow != null && firstRow != null) {
            headerRowColumnsCount = headerRow.getLastCellNum();
        }
    }
    private String convertDataToString(Cell cell){
        String sOut = "";
        if (cell != null) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    sOut = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    sOut = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    sOut = Double.toString(cell.getNumericCellValue());
                    break;
            }
        }
        return sOut;
    }
    private ColumnInfo getColumnInfo(int i) {
        ColumnInfo columnInfo = new ColumnInfo();
        Cell currentHeaderCell = headerRow.getCell(i);
        Cell currentFirstRowCell = firstRow.getCell(i);
        columnInfo.value = convertDataToString(currentHeaderCell);
        columnInfo.typeName = currentFirstRowCell != null ? currentFirstRowCell.getCellTypeEnum().toString() : "";
        columnInfo.columnNumber = i + 1;
        return columnInfo;
    }



    public String name;
    public ExcelWorksheet(Sheet sheet){
        this.sheet = sheet;
        this.name = sheet.getSheetName();
    }
    public ExcelWorksheet getData(){
        initializeSheet();
        if (headerRow != null && firstRow != null) {
            for (int i = 0; i < headerRowColumnsCount; i++) {
                ColumnInfo currentResult = getColumnInfo(i);
                results.add(currentResult);
            }
        }
        return this;
    }
    public ArrayList<ColumnInfo> getResults() {
        return results;
    }

}
