package com.wxystatic.officefilereaderlibrary;

import java.util.List;

/**
 * Created by static on 2017/11/29/029.
 */

public class ExcelResult {
    private Exception exception;
    public ExcelResult(){}
    public ExcelResult(Exception exception, List<ExcelModel> excelModelList) {
        this.exception = exception;
        this.excelModelList = excelModelList;
    }

    private List<ExcelModel> excelModelList;

    public Exception getException() {
        return exception;
    }

    public void setException(Exception exception) {
        this.exception = exception;
    }

    public List<ExcelModel> getExcelModelList() {
        return excelModelList;
    }

    public void setExcelModelList(List<ExcelModel> excelModelList) {
        this.excelModelList = excelModelList;
    }
}
