package com.wxystatic.officefilereaderlibrary;

import android.content.Context;
import android.icu.text.UFormat;
import android.os.AsyncTask;
import android.util.Log;
import android.widget.Toast;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * Created by static on 2017/11/29/029.
 */

public class ExcelReaderThread extends AsyncTask<String,Integer,ExcelResult>{
    private String path;
    private ExcelReadListener excelReadListener;
    public ExcelReaderThread(String path,ExcelReadListener excelReadListener){
        this.path=path;
        this.excelReadListener=excelReadListener;
    }
    @Override
    protected ExcelResult doInBackground(String... strings) {
        return readExcel(path);
    }
    public ExcelResult readExcel(String path){
        ExcelResult excelResult=new ExcelResult();
        List<ExcelModel> list=new ArrayList<>();
        excelResult.setExcelModelList(list);
        try {
            File file = new File(path);
            InputStream inputStream = new FileInputStream(file);
            Workbook workbook=Workbook.getWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(0);
            //行数
            int rowCount = sheet.getRows();
            Log.v("rowCount==",rowCount+"");
            if(rowCount>0){
                Cell cell = null;
            for (int row = 0; row < rowCount; row++){
                Cell[] cells=sheet.getRow(row);
                if (cells.length>0){
                    for (int column = 0; column < cells.length; column++){
                        cell = cells[column];
                        ExcelModel excelModel=new ExcelModel(row,column,cell.getContents().trim());
                        list.add(excelModel);
                    }
                }
            }
            }
            inputStream.close();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
            excelResult.setException(e);
        }
        return excelResult;
    }

    @Override
    protected void onPostExecute(ExcelResult excelResult) {
        if (excelResult.getException()==null){
            excelReadListener.onSuccess(excelResult.getExcelModelList());
        }else{
            excelReadListener.onFailed(excelResult.getException());
        }
    }
}
