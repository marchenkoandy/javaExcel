package com.company;

import com.company.excel.*;
import com.company.excel.thread.JThread;
import com.company.files.FileBrowser;
import com.company.records.*;
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
    private static final int                iMaxThreadsCount        = Runtime.getRuntime().availableProcessors();
    public static int                       LEVEL = 1;
    public static final int                 DEBUG_LEVEL = -1;
    private static final boolean            printDebug = false;

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
    private static void                     getAndPrintResultsAndCreateReport( String reportFile){
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
    private static void                     collectDataFromExcelFilesWithThreads(String path, String reportFile) {
        FileBrowser fb = new FileBrowser();
        fb.getFilesFromSingleFolder(new File(path));
        filesCount = fb.recursiveListOfFiles().size();
        printOnLevel(DEBUG_LEVEL,"Files count: " +filesCount);
        Vector<JThread>vectorJThread = new Vector<JThread>();
        int i=0;
        while (i<filesCount || vectorJThread.size()!=0) {
            if (i<filesCount && vectorJThread.size()<=iMaxThreadsCount) {
                String currentThreadName = "Thread_" + i;
                JThread thread = new JThread(currentThreadName, fb.recursiveListOfFiles().get(i));
                thread.start();
                vectorJThread.add(thread);
                i++;
            }
            else {
                for (int j=0;j<vectorJThread.size();j++) {
                    if (vectorJThread.get(j).isAlive()){
                        try {
                            vectorJThread.get(j).join();
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }
                    }
                    vectorJThread.remove(j);
                }
            }
        }
        getAndPrintResultsAndCreateReport(reportFile);
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
        getAndPrintResultsAndCreateReport(reportFile);
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
//        String path             = "C:/Users/amarchenko/Desktop/Java_Excel/vbs_password_1 - Copy";
//        String path             = "C:/Users/amarchenko/Desktop/Java_Excel/vbs_password_1/Dragon - Converted/Audio_Preservation/Audio_Preservation_After_CorrectMenu_SpellWindow - Copy.xlsx";
//        String reportFile       = "report.xlsx";
//        String reportFile       = "C:/Users/amarchenko/Desktop/Java_Excel/report.xlsx";
        Config config           = new Config(args);
        String actions          = config.getAction();
        String path             = config.getPath();
        String reportFile       = config.getResultFile();

        int timesToRepeat = 1;
        long sum = 0;
        Vector<Long> repeatTimes = new Vector<Long>();
        for (int i=1;i<=timesToRepeat;i++) {

            long start = System.currentTimeMillis();
            if (actions.equals(Actions.GET.toString())) {
                collectDataFromExcelFiles(path, reportFile);
            }
            else if (actions.equals(Actions.SET.toString())) {
                performExcelFilesUpdate(path, reportFile);
            }
            else if (actions.equals(Actions.COLLECT_DATA_FROM_EXCEL_FILES.toString())) {
//                collectDataFromExcelFiles(path, reportFile);
            }
            else if (actions.equals(Actions.COLLECT_DATA_FROM_EXCEL_FILES_WITH_THREADS.toString())) {
//                collectDataFromExcelFilesWithThreads(path, reportFile);
            }
            else if (actions.equals(Actions.PERFORM_EXCEL_FILES_UPDATE.toString())) {
//                performExcelFilesUpdate(path, reportFile);
            }
            long finish = System.currentTimeMillis();
            long duration = (finish - start);
            repeatTimes.add(duration);
            sum += duration;
        }

        long average = sum / repeatTimes.size();
        printOnLevel(LEVEL, "Files count is                         " + filesCount);
        printOnLevel(LEVEL, "Full execution time:\n                 " + sum);
        printOnLevel(DEBUG_LEVEL, "Average execution time:\n              " + average);
        for (int i=0;i<repeatTimes.size();i++){
            printOnLevel(DEBUG_LEVEL, repeatTimes.get(i).toString());
        }

    }
}
