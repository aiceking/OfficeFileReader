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
import java.util.Map;

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
   public void readExcel( String path,ExcelReadListener excelReadListener){
     new ExcelReaderThread(path,excelReadListener).execute();
   }
}
