package com.company.records;

import com.company.excel.Actions;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class Config {
    public                              Config(String[] args){
        fillClassFieldsWithArgs(args);
    }
    private void                        fillClassFieldsWithArgs(String[] args){
        Map<String,String> argsMap = new HashMap<String, String>();
        for(String item:args){
            String [] ar = item.split(":",2);
            argsMap.put(ar[0],ar[1]);
        }
        setAction                     (argsMap.get("action"));
        setPath                       (argsMap.get("path"));
        setResultFile                 (argsMap.get("resultFile"));
    }
    private String                      action;
    private String                      path;
    private String                      resultFile;

    public String                       getAction() {
        return action;
    }
    public void                         setAction(String action) {
        if (action.equals(Actions.GET.toString())
            ||action.equals(Actions.SET.toString())){
            this.action = action;
        }
        else{
            System.out.printf("Action '%s' is not supported.",action);
            System.exit(1);
        }
    }

    public String                       getResultFile() {
        return resultFile;
    }
    public void                         setResultFile(String resultFile) {
        if (getAction().equals(Actions.SET.toString())){
            File file = new File(resultFile);
            if (!file.exists()){
                System.out.printf("File '%s' must exist for action '%s'.",resultFile,getAction());
                System.exit(1);
            }
        }
        this.resultFile = resultFile;
    }

    public String                       getPath() {
        return path;
    }
    public void                         setPath(String path) {
        File file = new File(path);
        if (!file.exists()){
            System.out.printf("File/Folser '%s' must exist.",resultFile);
            System.exit(1);
        }
        this.path = path;
    }
}
