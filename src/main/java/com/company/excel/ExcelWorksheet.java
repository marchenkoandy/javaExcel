package com.company.excel;

import com.company.records.ColumnInfo;
import com.company.records.SheetInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ExcelWorksheet {
    private Row             headerRow;
    private Row             firstRow;
    private Sheet           sheet;
    private int             headerRowColumnsCount;

    public SheetInfo        sheetInfo;

    public static String    convertDataToString(Cell cell){
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
    private ColumnInfo      getColumnInfo(int i) {
        ColumnInfo columnInfo       = new ColumnInfo();
        Cell currentHeaderCell      = headerRow.getCell(i);
        Cell currentFirstRowCell    = firstRow.getCell(i);
        columnInfo.value            = convertDataToString(currentHeaderCell);
        columnInfo.type             = currentFirstRowCell != null ? currentFirstRowCell.getCellTypeEnum() : CellType.BLANK;
        columnInfo.columnNumber     = i + 1;
        return columnInfo;
    }

    public                  ExcelWorksheet(Sheet sheet){
        this.sheet                  = sheet;
        headerRow                   = sheet.getRow(0);
        firstRow                    = sheet.getRow(1);
        if (headerRow != null && firstRow != null) {
            headerRowColumnsCount   = headerRow.getLastCellNum();
        }
        sheetInfo                   = new SheetInfo();
        sheetInfo.columnInfos       = new ArrayList<ColumnInfo>();
    }
    public SheetInfo        getData(){
        if (headerRow != null && firstRow != null) {
            for (int i = 0; i < headerRowColumnsCount; i++) {
                sheetInfo.columnInfos.add(getColumnInfo(i));
            }
        }
        sheetInfo.name = this.sheet.getSheetName();
        return sheetInfo;
    }
}
