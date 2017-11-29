package com.wxystatic.officefilereaderlibrary;

import java.util.List;

/**
 * Created by static on 2017/11/29/029.
 */

public interface ExcelReadListener {
    void onSuccess(List<ExcelModel> list);
    void onFailed(Exception e);
}
