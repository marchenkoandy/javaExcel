package com.company.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWorkbook {
    private final String xlsx = ".xlsx";
    private final String xls = ".xls";
    private String inputFile;
    private ArrayList<Result> currentResults;
    public ExcelWorkbook setInputFile(String inputFile) {
        this.inputFile = inputFile;
        return this;
    }
    public ExcelWorkbook read() throws IOException {
        currentResults = new ArrayList<Result>();
        File file = new File(inputFile);
        FileInputStream inputStream = new FileInputStream(inputFile);
        Workbook currentWorkbook=null;
        String fileName = file.getName();
        String fileExtentionName = fileName.substring(fileName.lastIndexOf("."));
        if (fileExtentionName.equals(xlsx)) {
            currentWorkbook=new XSSFWorkbook (inputStream);
        }
        else if (fileExtentionName.equals(xls)) {
            currentWorkbook=new HSSFWorkbook(inputStream);
        }

        try {
            for (int currentSheetNumber=0;currentSheetNumber<= currentWorkbook.getNumberOfSheets()-1;currentSheetNumber++) {
                Sheet currentSheet = currentWorkbook.getSheetAt(currentSheetNumber);
                Row headerRow = currentSheet.getRow(0);
                int headerColumnsCount = headerRow.getLastCellNum();
                int newRowsCount = headerColumnsCount;
                for (int currentRowNumber=1;currentRowNumber<=currentSheet.getLastRowNum();currentRowNumber++) {
                    Row currentDataRow = currentSheet.getRow(currentRowNumber);
                    int currenrRowColumnsCount = currentDataRow.getLastCellNum();
                    if (newRowsCount < currenrRowColumnsCount){
                        newRowsCount = currenrRowColumnsCount;
                        System.out.println("Invalid columns count");
                    }
                }
                for (int i=0;i<=newRowsCount-1;i++){
                    Cell currentHeaderCell          = headerRow.getCell(i);
                    Cell currentDataCell            = currentSheet.getRow(1).getCell(i);

                    Result currentResult            = new Result();
                    currentResult.cellValue         = i<=headerColumnsCount-1?currentHeaderCell.getStringCellValue():"";
                    currentResult.cellTypeName      = currentDataCell.getCellTypeEnum().toString();
                    currentResult.cellSheetName     = currentSheet.getSheetName();
                    currentResult.cellWorkbookName  = inputFile;
                    currentResult.cellColumnNumber  = i + 1;
                    currentResults.add(currentResult);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return this;
    }
    public void print(){
        for (Result r:currentResults) {
            String line;
            line = "%s  %s  %s  %s  %s";
            line = String.format(line,r.cellValue,r.cellTypeName,r.cellSheetName,r.cellWorkbookName,r.cellColumnNumber);
            System.out.println(line);
        }
    }
}
