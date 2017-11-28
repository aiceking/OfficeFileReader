package com.wxystatic.officefilereaderlibrary;

import android.content.Context;
import android.widget.Toast;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * Created by static on 2017/11/28/028.
 */

public class ExcelReaderHelp {
    private static ExcelReaderHelp excelReaderHelp;
    public static ExcelReaderHelp getInstance(){
        if (excelReaderHelp==null){
            synchronized (ExcelReaderHelp.class){
                if (excelReaderHelp==null){
                    excelReaderHelp=new ExcelReaderHelp();
                }
            }
        }
        return excelReaderHelp;
    }
   public List<ExcelModel> readExcel(Context context, InputStream inputStream){
        List<ExcelModel> list=new ArrayList<>();
       try {
           Workbook workbook=Workbook.getWorkbook(inputStream);
           Sheet sheet = workbook.getSheet(0);
           //列数
           int columnCount = sheet.getColumns();
           //行数
           int rowCount = sheet.getRows();
           if(columnCount>0&&rowCount>0){
               //单元格
               Cell cell = null;
               for (int row = 0; row < rowCount; row++) {
                   for (int column = 0; column < columnCount; column++) {
                       cell = sheet.getCell(column, row);
                       ExcelModel excelModel=new ExcelModel(row,column,cell.getContents().trim());
                       list.add(excelModel);
                   }
               }
           }
           //关闭workbook,防止内存泄露
           workbook.close();
       } catch (FileNotFoundException e) {
           e.printStackTrace();

       } catch (IOException e) {
           Toast.makeText(context, "IO异常", Toast.LENGTH_SHORT).show();
           e.printStackTrace();
       } catch (BiffException e) {
           Toast.makeText(context, "文件格式异常", Toast.LENGTH_SHORT).show();
           e.printStackTrace();
       }
       return list;
   }
}
