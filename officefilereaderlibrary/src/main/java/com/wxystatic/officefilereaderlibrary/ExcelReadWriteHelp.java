package com.wxystatic.officefilereaderlibrary;

/**
 * Created by static on 2017/11/28/028.
 */

public class ExcelReadWriteHelp {
    private static ExcelReadWriteHelp excelReadWriteHelp;
    public static ExcelReadWriteHelp getInstance(){
        if (excelReadWriteHelp ==null){
            synchronized (ExcelReadWriteHelp.class){
                if (excelReadWriteHelp ==null){
                    excelReadWriteHelp =new ExcelReadWriteHelp();
                }
            }
        }
        return excelReadWriteHelp;
    }
   public void readExcel( String path,ExcelReadListener excelReadListener){
     new ExcelReaderThread(path,excelReadListener).execute();
   }
    public void WriteExcel( String path,String id,String columnName,String content,ExcelWriteListener excelWriteListener){
        new ExcelWriteThread(path,id,columnName,content,excelWriteListener).execute();
    }
}
