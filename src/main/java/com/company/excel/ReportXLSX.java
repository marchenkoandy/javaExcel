package com.company.excel;

import com.company.Main;
import com.company.records.Result;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.awt.*;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by AMarchenko on 7/6/2017.
 */
public class ReportXLSX  extends WorkbookBase {
    private final String                        RESULTS_ALL        = "RESULTS_ALL";
    private final String                        RESULTS_UNIQUE     = "RESULTS_UNIQUE";
    public                                      ReportXLSX(File fileExcel) {
        super(fileExcel);
    }

    private void                                deleteSheet(String sheetName){
        if (workbook.getSheet(sheetName) != null) {
            workbook.removeSheetAt(workbook.getSheetIndex(sheetName));
        }
    }
    private void                                reCreateSheet(String sheetName) {
        this.deleteSheet(sheetName);
        workbook.createSheet(sheetName);
    }
    private void                                setSingleRow(Sheet sheet, Point p, Result r){
        Row row = sheet.createRow(p.y);
        row.createCell(p.x + 0).setCellValue(r.cellValue);
        row.createCell(p.x + 1).setCellValue(r.cellType.toString());
        row.createCell(p.x + 2).setCellValue(r.cellSheetName);
        row.createCell(p.x + 3).setCellValue(r.cellWorkbookName);
    }
    public void                                 setData(ArrayList<Result> results){
        get();
        reCreateSheet(RESULTS_ALL);
        for (int i=0;i<results.size();i++) {
            Point p = new Point();
            p.setLocation(0,i);
            setSingleRow(workbook.getSheet(RESULTS_ALL),p, results.get(i));
        }
    }
    public void                                 setData(HashMap<String,Result> uniqueResults ){
        get();
        reCreateSheet(RESULTS_UNIQUE);
        int i=0;
        for (Map.Entry<String,Result> item: uniqueResults.entrySet()) {
            Point p = new Point();
            p.setLocation(1,i);
            setSingleRow(workbook.getSheet(RESULTS_UNIQUE), p, item.getValue());
            i++;
        }
    }
    public HashMap<String,String>               getData() {
        get();
        Main.printOnLevel(Main.DEBUG_LEVEL,"GETTING DATA...: " + name);
        HashMap<String, String> uniqueHeader = new HashMap<String, String>();
        Sheet sheet = workbook.getSheet(RESULTS_UNIQUE);
        if (sheet != null) {
            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String key      = this.getStringDataValue(row.getCell(1));
                String value    = this.getStringDataValue(row.getCell(0));
                boolean b1      = !key.equals("");
                boolean b2      = !value.equals("");
                boolean b3      = !uniqueHeader.containsKey(key);
                if (b1 && b2 && b3){
                    uniqueHeader.put(key, value);
                }
            }
        }
        this.close();
        return uniqueHeader;
    }
}
