package com.github.coorr.excel.core;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.util.WorkbookUtil;
import org.springframework.stereotype.Service;

import java.util.List;


@Service
@RequiredArgsConstructor
public class ExcelDownloadHandler {
    public <T> BaseWorkBook createWorkBook(Class<T> type, String sheetName) {
        return new BaseWorkBook<>(type, WorkbookUtil.createSafeSheetName(sheetName), null,this);
    }

    public <T> BaseWorkBook createWorkBook(Class<T> type, String sheetName, List<String> drawColumnList) {
        return new BaseWorkBook<>(type, WorkbookUtil.createSafeSheetName(sheetName), drawColumnList,this);
    }
}
