package com.company.excel;

import com.company.Main;
import com.company.records.SheetInfo;
import com.company.records.WorkbookInfo;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by AMarchenko on 7/5/2017.
 */
public class ExcelWorkbook extends WorkbookBase{

    public WorkbookInfo         workbookInfo;
    public void                 setNeedToBeSaved(boolean needToBeSaved) {
        if (needToBeSaved != false){
            this.needToBeSaved = needToBeSaved;
        }
    }

    public boolean              isNeedToBeSaved() {
        return needToBeSaved;
    }

    private boolean             needToBeSaved;
    public                      ExcelWorkbook(File fileExcel) {
        super(fileExcel);
        workbookInfo                = new WorkbookInfo();
        workbookInfo.sheetInfos     = new ArrayList<SheetInfo>();
    }
    public void                 getData(){
        Main.printOnLevel(Main.DEBUG_LEVEL,"GETTING DATA...: " + name);
        for (int i = 0; i< workbook.getNumberOfSheets(); i++) {
            ExcelWorksheet worksheet = new ExcelWorksheet(workbook.getSheetAt(i));
            workbookInfo.sheetInfos.add(worksheet.getData());
        }
        workbookInfo.name = this.name;
        this.close();
    }
    public void                 setData(HashMap<String, String> data){
        Main.printOnLevel(Main.DEBUG_LEVEL,"SETTING DATA...: " + name);
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
             ExcelWorksheet worksheet = new ExcelWorksheet(workbook.getSheetAt(i));
             worksheet.setData(data);
             setNeedToBeSaved(worksheet.isNeedToBeSaved());
        }
    }
}
