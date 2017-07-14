package com.company;

import com.company.excel.*;
import com.company.excel.thread.JThread;
import com.company.files.FileBrowser;
import com.company.records.ColumnInfo;
import com.company.records.Result;
import com.company.records.SheetInfo;
import com.company.records.WorkbookInfo;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.util.*;

public class Main {

    public static synchronized void setWorkbookInfos(ArrayList<WorkbookInfo> toSet) {


    }


    public static synchronized ArrayList<WorkbookInfo> getWorkbookInfos() {
        return workbookInfos;
    }

    private static ArrayList<WorkbookInfo>  workbookInfos           = new ArrayList<WorkbookInfo>();
    private static ArrayList<Result>        results                 = new ArrayList<Result>();
    private static HashMap<String,Result>   uniqueResults           = new HashMap<String, Result>();

    private static int                      filesCount;
    public static int                       LEVEL = 1;
    public static final int                 DEBUG_LEVEL = -1;
    private static final boolean            printDebug = true;

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
    private static void                     info(){
        printOnLevel(LEVEL, delimiter());
        printOnLevel(LEVEL,"Files count is: " + filesCount);
        printOnLevel(LEVEL, delimiter());
    }

    private static ArrayList<Result>        getAllResults(){
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
    private static HashMap<String,Result>   getUniqueResults(){
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
    private static void                     printUniqueResults(CellType type){
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
    private static void                     collectDataFromExcelFilesWithTheads(String path, String reportFile) {
        FileBrowser fb = new FileBrowser();
        fb.getFilesFromSingleFolder(new File(path));
        filesCount = fb.recursiveListOfFiles().size();
        printOnLevel(DEBUG_LEVEL,"Files count: " +filesCount);
        for (int i=0;i<filesCount-5;i=i+5) {
            JThread t1 = new JThread("t1 " + i+0, fb.recursiveListOfFiles().get(i+0));
            t1.start();
            JThread t2 = new JThread("t2 " + i+1, fb.recursiveListOfFiles().get(i+1));
            t2.start();
            JThread t3 = new JThread("t3 " + i+2, fb.recursiveListOfFiles().get(i+2));
            t3.start();
            JThread t4 = new JThread("t4 " + i+3, fb.recursiveListOfFiles().get(i+3));
            t4.start();
            JThread t5 = new JThread("t5 " + i+4, fb.recursiveListOfFiles().get(i+4));
            t5.start();
            try {
                t1.join();
                t2.join();
                t3.join();
                t4.join();
                t5.join();
            } catch (InterruptedException e) {
                e.printStackTrace();
            }


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
//        reportXLSX.setData(results);
        reportXLSX.setData(uniqueResults);
        reportXLSX.write();
        printOnLevel(DEBUG_LEVEL,"Were run " + JThread.counter + " threads..");
    }
    private static void                     collectDataFromExcelFiles(String path, String reportFile){
        FileBrowser fb = new FileBrowser();
        fb.getFilesFromSingleFolder(new File(path));
        filesCount = fb.recursiveListOfFiles().size();
        for (String file : fb.recursiveListOfFiles()) {
            ExcelWorkbook excelWorkbook = new ExcelWorkbook(new File(file));
            excelWorkbook.read();
            excelWorkbook.getData();
            excelWorkbook.close();
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
//        reportXLSX.setData(results);
        reportXLSX.setData(uniqueResults);
        reportXLSX.write();
    }
    private static void                     performExcelFilesUpdate(String path,String reportFile){
        File reportFileXLS = new File(reportFile);
        if (reportFileXLS.exists()) {
            ReportXLSX reportXLSX = new ReportXLSX(reportFileXLS);
            HashMap<String, String> uniqueResults = reportXLSX.getData();
            if (uniqueResults.size()>0) {
                int changedFilesCount = 0;
                FileBrowser fb = new FileBrowser();
                fb.getFilesFromSingleFolder(new File(path));
                filesCount = fb.recursiveListOfFiles().size();
                for (String file : fb.recursiveListOfFiles()) {
                    ExcelWorkbook excelWorkbook = new ExcelWorkbook(new File(file));
                    excelWorkbook.read();
                    excelWorkbook.setData(uniqueResults);
                    if (excelWorkbook.isNeedToBeSaved()) {
                        excelWorkbook.write();
                        changedFilesCount ++;
                        printOnLevel(LEVEL,"FILE was changed...: " + file);
                    }
                    excelWorkbook.close();
                }
                if (changedFilesCount==0){
                    printOnLevel(LEVEL,"No files of " + filesCount + " were changed.");
                }
                else {
                    printOnLevel(LEVEL,changedFilesCount +" file(s) of " + filesCount + " was(were) changed.");
                }
                printOnLevel(DEBUG_LEVEL, "Files count: " + filesCount);
            }
        }
    }

    public static void                      main(String[] args){
//        String path             = "C:/Users/user/tmp/Java_Excel";
//        String reportFile       = "C:/Users/user/tmp/report.xlsx";
        String path             = "C:/Users/amarchenko/Desktop/Java_Excel/vbs_password_1 - Copy";
//        String path             = "C:/Users/amarchenko/Desktop/Java_Excel/vbs_password_1/Dragon - Converted/Audio_Preservation/Audio_Preservation_After_CorrectMenu_SpellWindow - Copy.xlsx";
        String reportFile       = "C:/Users/amarchenko/Desktop/Java_Excel/report.xlsx";
        Actions actions         = Actions.COLLECT_DATA_FROM_EXCEL_FILES_WITH_THREADS;
//        Thread t = Thread.currentThread(); // получаем главный поток
//        System.out.println(t.getName()); // main
        long start = System.currentTimeMillis();
        switch (actions) {
            case COLLECT_DATA_FROM_EXCEL_FILES:
                collectDataFromExcelFiles(path, reportFile);
                break;
            case COLLECT_DATA_FROM_EXCEL_FILES_WITH_THREADS:
                collectDataFromExcelFilesWithTheads(path, reportFile);
                break;
            case PERFORM_EXCEL_FILES_UPDATE:
                performExcelFilesUpdate(path,reportFile);
                break;
        }
        long finish = System.currentTimeMillis();
        printOnLevel(DEBUG_LEVEL,"Code was executed in: " + (finish - start));
    }


}
