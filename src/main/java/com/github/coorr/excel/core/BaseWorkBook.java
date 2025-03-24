package com.github.coorr.excel.core;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class BaseWorkBook<T> extends AbstractExcelSxssfWorkBook<T> implements AutoCloseable {
    @Value("${content.type.version.2007}")
    private static String excelVersion2007;
    private final ExcelDownloadHandler handler;

    protected BaseWorkBook(Class<T> type, String sheetName, List<String> drawColumnList, ExcelDownloadHandler handler) {
        super(type, sheetName, drawColumnList);
        this.handler = handler;
    }

    @Override
    public void drawBody(T contents) {
        super.drawBody(contents);
    }

    public void output(HttpServletResponse response, String fileName) throws IOException {
        response.setContentType(excelVersion2007);
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName);

        try (OutputStream fileOut = response.getOutputStream()) {
            getWorkbook().write(fileOut);
        }

        response.getOutputStream().flush();
        response.getOutputStream().close();
    }

    @Override
    public void close() throws Exception {
        getWorkbook().close();
    }
}
