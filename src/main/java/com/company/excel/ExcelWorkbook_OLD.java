package com.company.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWorkbook_OLD {
    public static final String XLSX = ".xlsx";
    public static final String XLS = ".xls";
    public static final String TILDA = "~$";

    private String inputFile;
    public ExcelWorkbook_OLD setInputFile(String inputFile) {
        this.inputFile = inputFile;
        return this;
    }

    private ArrayList<Result> results = new ArrayList<Result>();
    public ArrayList<Result> getResults() {
        return results;
    }

    private Sheet currentSheet;
    private Row headerRow;
    private Row firstRow;
    private int headerRowColumnsCount;

    private void initializeSheet(){
//        System.out.println(currentSheet.getSheetName());
        headerRow = currentSheet.getRow(0);
        firstRow = currentSheet.getRow(1);
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
    private Result getColumnInfo(int i) {
        Result result = new Result();
        Cell currentHeaderCell = headerRow.getCell(i);
        Cell currentFirstRowCell = firstRow.getCell(i);
        result.cellValue = convertDataToString(currentHeaderCell);          //currentHeaderCell != null ? currentHeaderCell.getStringCellValue() : "";
        result.cellTypeName = currentFirstRowCell != null ? currentFirstRowCell.getCellTypeEnum().toString() : "";
        result.cellSheetName = currentSheet.getSheetName();
        result.cellWorkbookName = inputFile;
        result.cellColumnNumber = i + 1;
        return result;
    }

    public ArrayList<Result> read(){
        File file = new File(inputFile);
        FileInputStream inputStream=null;
        try{
            inputStream= new FileInputStream(inputFile);

        Workbook currentWorkbook=null;
        String fileName = file.getName();
        String fileExtensionName = fileName.substring(fileName.lastIndexOf("."));
        if (fileExtensionName.equals(XLSX)) {
            currentWorkbook=new XSSFWorkbook (inputStream);
        }
        else if (fileExtensionName.equals(XLS)) {
            currentWorkbook=new HSSFWorkbook(inputStream);
        }
            for (int currentSheetNumber=0;currentSheetNumber<currentWorkbook.getNumberOfSheets();currentSheetNumber++) {
                currentSheet = currentWorkbook.getSheetAt(currentSheetNumber);
                initializeSheet();
                if (headerRow != null && firstRow != null) {
                    for (int i = 0; i < headerRowColumnsCount; i++) {
                        Result currentResult = getColumnInfo(i);
                        results.add(currentResult);
                    }
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            if (inputStream != null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return results;
    }
    private static void printSingleRecord(Result r){
        String line;
        line = "%s %s %s %s %s";
        line = String.format(line,r.cellValue,r.cellTypeName,r.cellSheetName,r.cellWorkbookName,r.cellColumnNumber);
        System.out.println(line);
    }
    public static void print(ArrayList<Result> results){
        for (Result r: results) {
            printSingleRecord(r);
        }
        System.out.println("Full amount of records is: " + results.size());
    }
    public static HashMap<String,Result> uniqueList(ArrayList<Result> results){
        HashMap <String,Result> uniqueList = new HashMap<String, Result>();
        for (Result r:results) {
            if (!uniqueList.containsKey(r.cellValue)){
                uniqueList.put(r.cellValue,r);
            }
        }
        return uniqueList;
    }
    public static void print( HashMap<String,Result> list){
        for (Map.Entry<String,Result> item: list.entrySet()) {
            if (!item.getValue().cellTypeName.equals("STRING")) {
//                System.out.println(item.getKey());
                printSingleRecord(item.getValue());
            }
        }
        System.out.println("Full amount of unique records is: " + list.size());
    }

}
