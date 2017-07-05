package com.company;

import com.company.excel.ExcelWorkbook_OLD;
import com.company.excel.ExcelWorkbook;
import com.company.excel.Result;
import com.company.files.FileBrowser;
import java.io.File;
import java.util.ArrayList;

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
    public static String delimeter(){
        return "========================================";
    }
    public static ArrayList<ExcelWorkbook> allWorkbooks = new ArrayList<ExcelWorkbook>();
    public static int filesCount;
    public static int LEVEL = 1;

    public static void printOnLevel(int level, String line){
        String tab = "\t";
        String tabs = "";
        for (int i=1;i<level;i++){
            tabs += tab;
        }
        System.out.println(tabs + line);
    }
    public static void info(){
        printOnLevel(LEVEL,delimeter());
        printOnLevel(LEVEL,"Files count is: " + filesCount);
        printOnLevel(LEVEL,delimeter());
    }
    public static void printResults(){
        info();
        for (ExcelWorkbook workbook:allWorkbooks) {
            workbook.printResults();
        }
    }

    public static void main(String[] args){

        FileBrowser fb = new FileBrowser();
        fb.getFilesFromSingleFolder(new File("C:/Users/amarchenko/Desktop/Java_Excel"));
        filesCount = fb.recursiveListofFiles().size();

//        ArrayList<ExcelWorkbook> allWorkbooks = new ArrayList<ExcelWorkbook>();
        for (String file :fb.recursiveListofFiles()) {
//            System.out.println("Working with file: " + file);
            ExcelWorkbook excelWorkbook = new ExcelWorkbook();
            excelWorkbook.read(new File(file));
            allWorkbooks.add(excelWorkbook);
        }
        printResults();
//        ExcelWorkbook_OLD.print(currentResults);
//        ExcelWorkbook_OLD.print(ExcelWorkbook_OLD.uniqueList(currentResults));
    }
}
