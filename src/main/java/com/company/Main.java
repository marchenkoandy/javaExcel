package com.company;

import com.company.excel.*;
import com.company.files.FileBrowser;
import com.company.records.ColumnInfo;
import com.company.records.Result;
import com.company.records.SheetInfo;
import com.company.records.WorkbookInfo;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class Main {

    private static ArrayList<WorkbookInfo>  workbookInfos = new ArrayList<WorkbookInfo>();
    private static ArrayList<Result>        results = new ArrayList<Result>();
    private static HashMap<String,Result>   uniqueResults           = new HashMap<String, Result>();

    public static int filesCount;
    public static int LEVEL = 1;
    public static final int DEBUG_LEVEL = -1;
    public static final boolean printDebug = false;

    public static String                    delimiter(){
        return "========================================";
    }
    public static void                      printOnLevel(int level, String line) {
        String tab = "\t";
        String tabs = "";
        for (int i = 1; i < level; i++) {
            tabs += tab;
        }
        if (level!=DEBUG_LEVEL || printDebug)  {
            System.out.println(tabs + line);
        }
    }
    public static void                      info(){
        printOnLevel(LEVEL, delimiter());
        printOnLevel(LEVEL,"Files count is: " + filesCount);
        printOnLevel(LEVEL, delimiter());
    }


    public static ArrayList<Result>         getAllResults(){
        for (WorkbookInfo workbook: workbookInfos) {
            for (SheetInfo sheet:workbook.sheetInfos) {
                for (ColumnInfo info:sheet.columnInfos) {
                    Result r =new Result();
                    r.cellValue = info.value;
                    r.cellType = info.type;
                    r.cellColumnNumber = info.columnNumber;
                    r.cellSheetName = sheet.name;
                    r.cellWorkbookName = workbook.name;
                    results.add(r);
                }
            }
        }
        return results;
    }
    public static HashMap<String,Result>    getUniqueResults(){
        if (workbookInfos.size()!= 0 && results.size() == 0){
            getAllResults();
        }
        for (Result r: results) {
            if (!uniqueResults.containsKey(r.cellValue)){
                uniqueResults.put(r.cellValue,r);
            }
        }
        return uniqueResults;
    }
    private static void                     printSingleRecord(Result r){
        String line;
        line = "%s %s %s %s %s";
        line = String.format(line,r.cellValue,r.cellType,r.cellSheetName,r.cellWorkbookName,r.cellColumnNumber);
        printOnLevel(DEBUG_LEVEL,line);
    }
    public static void                      printResults(){
        info();
        for (WorkbookInfo workbook: workbookInfos) {
            workbook.printResults();
        }
    }
    public static void                      printAllResultRecords(){
        for (Result r: results) {
            printSingleRecord(r);
        }
        printOnLevel(LEVEL,"Full amount of records is: " + results.size());
    }
    public static void                      printUniqueResults(CellType type){
        int counter = 0;
        for (Map.Entry<String,Result> item: uniqueResults.entrySet()) {
            if (type == null || item.getValue().cellType == type){
                printSingleRecord(item.getValue());
                counter += 1;
            }
        }
        String delta = type != null? " with type '" + type.toString() + "' ": "";
        printOnLevel(DEBUG_LEVEL,"Full amount of unique records" + delta+": " + counter);
    }
    public static void                      collectDataFromExcelFiles(String path, String reportFile){
        FileBrowser fb = new FileBrowser();
        fb.getFilesFromSingleFolder(new File(path));
        filesCount = fb.recursiveListOfFiles().size();
        for (String file : fb.recursiveListOfFiles()) {
            ExcelWorkbook excelWorkbook = new ExcelWorkbook(new File(file));
            excelWorkbook.read();
            excelWorkbook.getData();
            workbookInfos.add(excelWorkbook.workbookInfo);
        }
        printOnLevel(DEBUG_LEVEL,"Files count: " +filesCount);
        getUniqueResults();

        printUniqueResults(null);
        printUniqueResults(CellType.BLANK);
        printUniqueResults(CellType._NONE);
        printUniqueResults(CellType.BOOLEAN);
        printUniqueResults(CellType.NUMERIC);
        printUniqueResults(CellType.ERROR);
        printUniqueResults(CellType.FORMULA);
        printUniqueResults(CellType.STRING);

        ReportXLSX reportXLSX = new ReportXLSX(new File(reportFile));
        reportXLSX.setUniqueData(uniqueResults);
        reportXLSX.write();
    }
    public static void                      performExcelFilesUpdate(String path,String reportFile){
        File file = new File(reportFile);
        if (file.exists()) {
            ReportXLSX reportXLSX = new ReportXLSX(file);
            reportXLSX.setData(results);
            HashMap<String, String> uniqueResults = reportXLSX.getData();
        }
    }

    public static void                      main(String[] args){
//        String path             = "C:/Users/user/tmp/Java_Excel";
//        String reportFile       = "C:/Users/user/tmp/report.xlsx";
        String path             = "C:/Users/amarchenko/Desktop/Java_Excel/vbs_password_1";
        String reportFile       = "C:/Users/amarchenko/Desktop/Java_Excel/report.xlsx";
        Actions actions         = Actions.COLLECT_DATA_FROM_EXCEL_FILES;
        switch (actions) {
            case COLLECT_DATA_FROM_EXCEL_FILES:
                collectDataFromExcelFiles(path, reportFile);
                break;
            case PERFORM_EXCEL_FILES_UPDATE:
                performExcelFilesUpdate(path,reportFile);
                break;
        }
    }
}
