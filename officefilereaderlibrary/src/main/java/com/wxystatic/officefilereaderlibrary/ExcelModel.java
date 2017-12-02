package com.wxystatic.officefilereaderlibrary;


/**
 * Created by static on 2017/11/28/028.
 */

public class ExcelModel {
    private int Row;
    private int Column;
    private String content;

    public ExcelModel(int row, int column, String content) {
        Row = row;
        Column = column;
        this.content = content;
    }

    public int getRow() {
        return Row;
    }

    public void setRow(int row) {
        Row = row;
    }

    public int getColumn() {
        return Column;
    }

    public void setColumn(int column) {
        Column = column;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }
}
