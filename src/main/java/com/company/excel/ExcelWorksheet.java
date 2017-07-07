package com.company.excel;

import com.company.records.ColumnInfo;
import com.company.records.SheetInfo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ExcelWorksheet {
    private Row                             headerRow;
    private Row                             firstRow;
    private Sheet                           sheet;
    private int                             headerRowColumnsCount;
    private HashMap<String, String>         data;



    public void setNeedToBeSaved(boolean needToBeSaved) {
        if (needToBeSaved != false){
            this.needToBeSaved = needToBeSaved;
        }
    }

    public boolean isNeedToBeSaved() {
        return needToBeSaved;
    }

    private boolean                         needToBeSaved;
    public SheetInfo                        sheetInfo;


    private ColumnInfo      getColumnInfo(int i) {
        ColumnInfo columnInfo       = new ColumnInfo();
        Cell currentHeaderCell      = headerRow.getCell(i);
        Cell currentFirstRowCell    = firstRow.getCell(i);
        columnInfo.value            = WorkbookBase.getStringDataValue(currentHeaderCell);
        columnInfo.type             = currentFirstRowCell != null ? currentFirstRowCell.getCellTypeEnum() : CellType.BLANK;
        columnInfo.columnNumber     = i + 1;
        return columnInfo;
    }
    private void            setColumnInfo(int i) {
        Cell currentHeaderCell      = headerRow.getCell(i);
        String actualValue          = WorkbookBase.getStringDataValue(currentHeaderCell);
        String associatedValue      = data.get(actualValue);
        boolean b1 = !actualValue.equals("");
        boolean b2 = (associatedValue != null && !associatedValue.equals(""));
        boolean b3 = !actualValue.equals(associatedValue);
        boolean b4 = b1 && b2 && b3;
        setNeedToBeSaved(b4);
        if (    b4   ){
            currentHeaderCell.setCellValue(associatedValue);
        }
    }
    public                  ExcelWorksheet(Sheet sheet){
        this.sheet                  = sheet;
        headerRow                   = sheet.getRow(0);
        firstRow                    = sheet.getRow(1);
        if (headerRow != null && firstRow != null) {
            headerRowColumnsCount   = headerRow.getLastCellNum();
        }
        setNeedToBeSaved(false);
    }
    public SheetInfo        getData(){
        sheetInfo                   = new SheetInfo();
        sheetInfo.columnInfos       = new ArrayList<ColumnInfo>();
        if (headerRow != null && firstRow != null) {
            for (int i = 0; i < headerRowColumnsCount; i++) {
                sheetInfo.columnInfos.add(getColumnInfo(i));
            }
        }
        sheetInfo.name = this.sheet.getSheetName();
        return sheetInfo;
    }

    public void             setData(HashMap<String, String> data) {
        this.data = data;
        if (headerRow != null && firstRow != null) {
            for (int i = 0; i < headerRowColumnsCount; i++) {
                setColumnInfo(i);
            }
        }
    }


}
