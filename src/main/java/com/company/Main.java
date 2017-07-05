package com.company;

import com.company.excel.*;
import com.company.files.FileBrowser;
import com.company.frame.MainDialog;
import org.apache.poi.ss.usermodel.CellType;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class Main {
    private void FormTest() {
//     Create a file chooser
//        MainDialog dialog = new MainDialog();
//        dialog.pack();
//        dialog.setVisible(true);
////        System.exit(0);
//        final JFileChooser fc = new JFileChooser();
//        FileNameExtensionFilter filter = new FileNameExtensionFilter("JPG & GIF Images", "jpg", "gif");
//        fc.setFileFilter(filter);
//        //In response to a button click:
//        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
//
//        fc.showOpenDialog(dialog);
//        int returnVal = fc.showOpenDialog();
//        ExcelWorkbook_OLD excelWorkbook = new ExcelWorkbook_OLD();
//        excelWorkbook.setInputFile("C:/Temp/ACO_LMO/ACO.XLSX").read().print();
//        MainDialog dialog = new MainDialog();
//        dialog.pack();
//        dialog.setVisible(true);
//        System.exit(0);
 }
    private void test(){
     FileBrowser fb = new FileBrowser();
     fb.getFilesFromSingleFolder(new File("C:/Users/amarchenko/Desktop/Java_Excel"));
     System.out.println("Files count is: " + fb.recursiveListofFiles().size());
     ArrayList<Result> currentResults = new ArrayList<Result>();
     for (String file :fb.recursiveListofFiles()) {
//            System.out.println("Working with file: " + file);
         ExcelWorkbook_OLD excelWorkbookOLD = new ExcelWorkbook_OLD();
         excelWorkbookOLD.setInputFile(file).read();
         currentResults.addAll(excelWorkbookOLD.getResults());
     }
//        ExcelWorkbook_OLD.print(currentResults);
     ExcelWorkbook_OLD.print(ExcelWorkbook_OLD.uniqueList(currentResults));
 }
    private static ArrayList<ExcelWorkbook> allWorkbooks = new ArrayList<ExcelWorkbook>();
    private static ArrayList<Result> allResults = new ArrayList<Result>();
    private static HashMap<String,Result> uniqueResults = new HashMap<String, Result>();

    public static int filesCount;
    public static int LEVEL = 1;

    public static String                    delimiter(){
        return "========================================";
    }
    public static void                      printOnLevel(int level, String line){
        String tab = "\t";
        String tabs = "";
        for (int i=1;i<level;i++){
            tabs += tab;
        }
        System.out.println(tabs + line);
    }
    public static void                      info(){
        printOnLevel(LEVEL, delimiter());
        printOnLevel(LEVEL,"Files count is: " + filesCount);
        printOnLevel(LEVEL, delimiter());
    }


    public static ArrayList<Result>         getAllResults(){
        for (ExcelWorkbook workbook:allWorkbooks) {
            for (ExcelWorksheet sheet:workbook.sheets) {
                for (ColumnInfo info:sheet.getColumnInfos()) {
                    Result r =new Result();
                    r.cellValue = info.value;
                    r.cellType = info.type;
                    r.cellColumnNumber = info.columnNumber;
                    r.cellSheetName = sheet.name;
                    r.cellWorkbookName = workbook.name;
                    allResults.add(r);
                }
            }
        }
        return allResults;
    }
    public static HashMap<String,Result>    getUniqueResults(){
        if (allWorkbooks.size()!= 0 && allResults.size() == 0){
            getAllResults();
        }
        for (Result r:allResults) {
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
        printOnLevel(LEVEL,line);
    }
    public static void                      printResults(){
        info();
        for (ExcelWorkbook workbook:allWorkbooks) {
            workbook.printResults();
        }
    }
    public static void                      printAllResultRecords(){
        for (Result r: allResults) {
            printSingleRecord(r);
        }
        printOnLevel(LEVEL,"Full amount of records is: " + allResults.size());
    }
    public static void                      printUniqueResults(CellType type){
        int counter = 0;
        for (Map.Entry<String,Result> item: uniqueResults.entrySet()) {
            if (type == null || item.getValue().cellType == type){
//                printSingleRecord(item.getValue());
                counter += 1;
            }
        }
        String delta = type != null? " with type '" + type.toString() + "' ": "";
        printOnLevel(LEVEL,"Full amount of unique records" + delta+": " + counter);
    }
    public static void                      main(String[] args) throws IOException {
        boolean bGUI = true;
        if (!bGUI) {
            FileBrowser fb = new FileBrowser();
            fb.getFilesFromSingleFolder(new File("C:/Users/amarchenko/Desktop/Java_Excel"));
            filesCount = fb.recursiveListofFiles().size();
            for (String file : fb.recursiveListofFiles()) {
                ExcelWorkbook excelWorkbook = new ExcelWorkbook();
                excelWorkbook.read(new File(file));
                allWorkbooks.add(excelWorkbook);
            }
//        printResults();
//        System.in.read();

//        getAllResults();
//        printAllResultRecords();


            getUniqueResults();


            printUniqueResults(null);
            printUniqueResults(CellType.BLANK);
            printUniqueResults(CellType._NONE);
            printUniqueResults(CellType.BOOLEAN);
            printUniqueResults(CellType.NUMERIC);
            printUniqueResults(CellType.ERROR);
            printUniqueResults(CellType.FORMULA);
            printUniqueResults(CellType.STRING);
        }
        else {
            MainDialog dialog = new MainDialog();
            dialog.setVisible(true);

//            System.exit(0);
        }
    }
}
