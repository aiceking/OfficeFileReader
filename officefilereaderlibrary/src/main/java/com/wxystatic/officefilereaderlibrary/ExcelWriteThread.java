package com.wxystatic.officefilereaderlibrary;

import android.os.AsyncTask;
import android.util.Log;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 * Created by static on 2017/11/29/029.
 */

public class ExcelWriteThread extends AsyncTask<String,Integer,ExcelResult>{
    private String path;
    private ExcelWriteListener excelWriteListener;
    private String id;
    private String columnName;
    private String content;
    public ExcelWriteThread(String path,String id,String columnName,String content, ExcelWriteListener excelWriteListener){
        this.path=path;
        this.id=id;
        this.columnName=columnName;
        this.content=content;
        this.columnName=columnName;
        this.excelWriteListener=excelWriteListener;
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

            Sheet sheet_read=workbook.getSheet(0);
            int rowCount =sheet_read.getRows();
            int columnCount =sheet_read.getColumns();
            int search_row=-1;
            int search_column=-1;
            if (rowCount>0&&columnCount>0){
                Cell cell_search=null;
                Cell[] cells_one=sheet_read.getRow(0);
            for (int row=0;row<rowCount;row++){
                for (int column=0;column<columnCount;column++){
                    cell_search=sheet_read.getCell(column,row);
                    if (cell_search.getContents().trim().equals(id)){
                        search_row=row;
                        for (int i=0;i<cells_one.length;i++){
                            if (cells_one[i].getContents().trim().equals(columnName)){
                                search_column=i;
                                break;
                            }
                        }
                    }
            }
            }
            if (search_row>0&&search_column>0){
                WritableWorkbook book= Workbook.createWorkbook(new File(path),workbook);
                WritableSheet sheet_write = book.getSheet(0);
                WritableCell cell_write = sheet_write.getWritableCell( search_column,search_row);
                Label label=new Label(search_column,search_row ,content);
                label.setCellFormat(cell_write.getCellFormat());
                sheet_write.addCell(label);
                book.write();
                book.close();
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
            excelWriteListener.onSuccess();
        }else{
            excelWriteListener.onFailed(excelResult.getException());
        }
    }
}
