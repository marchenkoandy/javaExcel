package com.company.excel.thread;

import com.company.Main;
import com.company.excel.ExcelWorkbook;

import java.io.File;

/**
 * Created by AMarchenko on 7/14/2017.
 */
public class JThread extends Thread{
    private String file;
    public static int counter;

    public JThread(String name,String file){
        super(name);
        this.file=file;
        counter++;

    }
    public void run(){
//        System.out.printf("Thread %s begins to work\n", Thread.currentThread().getName());
        try{
            ExcelWorkbook excelWorkbook = new ExcelWorkbook(new File(file));
            excelWorkbook.read();
            excelWorkbook.getData();
            excelWorkbook.close();
            Main.getWorkbookInfos().add(excelWorkbook.workbookInfo);
        }
        catch (Exception e){
            Main.printOnLevel(Main.DEBUG_LEVEL,e.getMessage());
            throw new RuntimeException();
        }
//        System.out.printf("Thread %s stops to work\n", Thread.currentThread().getName());
    }
}
