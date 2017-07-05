package com.company.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ExcelWorkbook_ToRefact {
    public static final String XLSX = ".xlsx";
    public static final String XLS = ".xls";
    public static final String TILDA = "~$";

    public String name;
    public ArrayList<ExcelWorksheet> sheets = new ArrayList<ExcelWorksheet>();


    public void read(File fileExcel){
        FileInputStream inputStream=null;
        this.name = fileExcel.getAbsolutePath();
        try{
            inputStream= new FileInputStream(fileExcel);
            Workbook currentWorkbook=null;
            String fileName = fileExcel.getName();
            String fileExtensionName = fileName.substring(fileName.lastIndexOf("."));
            if (fileExtensionName.equals(XLSX)) {
                currentWorkbook=new XSSFWorkbook(inputStream);
            }
            else if (fileExtensionName.equals(XLS)) {
                currentWorkbook=new HSSFWorkbook(inputStream);
            }
            for (int i=0;i<currentWorkbook.getNumberOfSheets();i++) {
                ExcelWorksheet worksheet = new ExcelWorksheet(currentWorkbook.getSheetAt(i));
                sheets.add(worksheet.getData());
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (inputStream != null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}