package com.github.coorr.excel.core;

import com.github.coorr.excel.annotation.ExcelColumn;
import com.github.coorr.excel.exception.ExcelFieldAccessException;
import com.github.coorr.excel.support.WorkBookUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.ObjectUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

@Slf4j
public abstract class AbstractExcelSxssfWorkBook<T> {
    private final int ROW_ACCESS_WINDOW_SIZE = 500;
    private final String EXCEL_BASIC_FONT_NAME = "Arial";
    private final int ROW_START_INDEX = 1;
    SpreadsheetVersion supplyExcelVersion2007 = SpreadsheetVersion.EXCEL2007;
    private final int xlsxMaxRow = supplyExcelVersion2007.getMaxRows() - 1; // 최대 행수

    private int bodyRowStartIndex = 1;
    private int bodyRowCount = 0;
    private final List<String> excelHeader = new ArrayList<>();

    private final Workbook workbook;
    private final CellStyle headerStyle;
    private final CellStyle contentsStyle;
    private Sheet sheet;
    private int sheetNumber = 0;

    private boolean hasPreviousMergedHeader;
    private String previousHeaderName;
    private int headerMergeStartIndex;

    private List<String> drawColumnList;
    private final List<Field> fields;

    protected AbstractExcelSxssfWorkBook(Class<T> type, String sheetName, List<String> drawColumnList) {
        this.workbook = new SXSSFWorkbook(new XSSFWorkbook(), ROW_ACCESS_WINDOW_SIZE, true);

        // 기본 폰트
        Font defaultFont = workbook.createFont();

        defaultFont.setFontName(EXCEL_BASIC_FONT_NAME);

        // 제목 스타일
        headerStyle = workbook.createCellStyle();
        this.makeHeaderStyle(headerStyle, defaultFont);

        // 본문 스타일
        contentsStyle = workbook.createCellStyle();
        this.makeContentsStyle(contentsStyle, defaultFont);

        this.sheet = workbook.createSheet(sheetName);

        this.fields = getAllFields(type);

        this.drawColumnList = drawColumnList;

        this.drawHeader();
    }


    protected void drawHeader() {
        Row firstRow = sheet.createRow(ROW_START_INDEX-1);
        Row secondRow = sheet.createRow(ROW_START_INDEX);
        this.renderHeader(firstRow, secondRow, headerStyle);
    }

    protected void drawBody(T contents) {
        if (!ObjectUtils.isEmpty(contents)) {
            this.renderBody(contents);
            bodyRowCount++;
        }
    }

    protected void renderHeader(Row firstRow, Row secondRow, CellStyle style) {
        int drawColumnSize = ObjectUtils.isEmpty(this.drawColumnList) ? 0 : this.fields.size() - this.drawColumnList.size();

        int skipCnt = 0;
        for (Field field : this.fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            String fieldName = StringUtils.isNotEmpty(annotation.columnName()) ? annotation.columnName() : field.getName();

            if (!annotation.required() && !ObjectUtils.isEmpty(this.drawColumnList) && !this.drawColumnList.contains(fieldName)) {
                skipCnt++;
                continue;
            }

            int columnIndex = (drawColumnSize > 0 && skipCnt > 0) ? annotation.column() - skipCnt : annotation.column();

            if (annotation.row() == firstRow.getRowNum()) {
                addHeader(annotation, columnIndex, firstRow, secondRow, style);
            } else {
                try {
                    firstRow.getCell(columnIndex).getStringCellValue();
                } catch (NullPointerException npe) {
                    addHeader(annotation, columnIndex, firstRow, secondRow, style);
                }
            }
        }
    }

    private void addHeader(ExcelColumn annotation, int columnIndex, Row firstRow, Row secondRow, CellStyle style) {
        Cell firstHeader = firstRow.createCell(columnIndex);
        setCellHeaderValue(annotation.headerName(), firstHeader);
        firstHeader.setCellStyle(style);

        Cell secondHeader = secondRow.createCell(columnIndex);
        setCellHeaderValue(annotation.secondHeaderName(), secondHeader);
        secondHeader.setCellStyle(style);

        mergeHeaders(annotation, columnIndex);
        sheet.setColumnWidth(columnIndex, annotation.width());

        if (StringUtils.isNotBlank(annotation.headerName()) || StringUtils.isNotBlank(annotation.secondHeaderName())) {
            excelHeader.add(WorkBookUtils.createHeaderValue(annotation.headerName(), annotation.secondHeaderName()));
        }
    }

    private void setCellHeaderValue(String headerName, Cell cell) {
        cell.setCellValue(headerName);
    }

    private void mergeHeaders(ExcelColumn annotation, int columnIndex) {
        if(StringUtils.isBlank(annotation.secondHeaderName())) {
            mergePreviousHeaders(columnIndex);
            sheet.addMergedRegion(new CellRangeAddress(0, 1, columnIndex, columnIndex));
        } else {
            mergeIfDifferentHeadersAppear(annotation, columnIndex);
            previousHeaderName = annotation.headerName();
            hasPreviousMergedHeader = true;
        }
    }

    private void mergePreviousHeaders(int columnIndex) {
        if(hasPreviousMergedHeader){ // 이전에 병합되야할 헤더들 머지
            sheet.addMergedRegion(new CellRangeAddress(0, 0, headerMergeStartIndex, columnIndex -1));
            hasPreviousMergedHeader = false;
            previousHeaderName = null;
        }
    }

    private void mergeIfDifferentHeadersAppear(ExcelColumn annotation, int columnIndex) {
        if (!annotation.headerName().equals(previousHeaderName)) {
            if (previousHeaderName != null) { // 새로운 헤더값을 만났을경우 이전 헤더까지 머지
                sheet.addMergedRegion(new CellRangeAddress(0, 0, headerMergeStartIndex, columnIndex - 1));
            } else {
                headerMergeStartIndex = columnIndex;
            }
        }
    }

    protected void renderBody(T contents) {
        if (bodyRowStartIndex == this.get2007MaxRow()) {
            bodyRowStartIndex = ROW_START_INDEX;
            this.createNewSheetWithHeader();
        }
        this.addRow(++bodyRowStartIndex, contents);
    }

    private void addRow(int rowIndex, T contents) {
        Row row = sheet.createRow(rowIndex);
        int drawColumnSize = ObjectUtils.isEmpty(this.drawColumnList) ? 0 : this.fields.size() - this.drawColumnList.size();

        int skipCnt = 0;
        for (Field field : this.fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            String fieldName = StringUtils.isNotEmpty(annotation.columnName()) ? annotation.columnName() : field.getName();

            if (!annotation.required() && !ObjectUtils.isEmpty(this.drawColumnList) && !this.drawColumnList.contains(fieldName)) {
                skipCnt++;
                continue;
            }
            int column = (drawColumnSize > 0 && skipCnt > 0) ? annotation.column() - skipCnt : annotation.column();

            Cell cell = row.createCell(column);

            try {
                field.setAccessible(true);
                Object cellValue = field.get(contents);
                this.renderCellValue(cell, cellValue);
            } catch (IllegalAccessException e) {
                throw new ExcelFieldAccessException("Cannot access field: " + field.getName(), e);
            }
        }
    }

    protected void renderCellValue(Cell cell, Object content) {
        if (content instanceof Number) {
            Number numberValue = (Number) content;
            cell.setCellValue(numberValue.doubleValue());
            cell.setCellStyle(contentsStyle);
            return;
        }
        cell.setCellValue(content == null ? "" : content.toString());
        cell.setCellStyle(contentsStyle);
    }

    private void createNewSheetWithHeader() {
        this.sheet = workbook.createSheet(sheet.getSheetName().concat("_").concat(Integer.toString(sheetNumber++)));
        this.drawHeader();
    }

    public List<Field> getAllFields(Class<T> clazz) {
        List<Field> fields = new ArrayList<>();
        for (Class<T> clazzInClasses : getAllClassesIncludingSuperClasses(clazz, true)) {
            List<Field> annotatedFields = Arrays.stream(clazzInClasses.getDeclaredFields())
                    .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                    .collect(Collectors.toList());
            fields.addAll(annotatedFields);
        }
        return fields;
    }

    private List<Class<T>> getAllClassesIncludingSuperClasses(Class<T> clazz, boolean fromSuper) {
        List<Class<T>> classes = new ArrayList<>();
        while (clazz != null) {
            classes.add(clazz);
            clazz = (Class<T>) clazz.getSuperclass();
        }
        if (fromSuper) {
            Collections.reverse(classes);
        }
        return classes;
    }

    protected void makeContentsStyle(CellStyle style, Font font) {
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
    }

    protected void makeHeaderStyle(CellStyle style, Font font) {
        style.setFont(font);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setWrapText(true);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public CellStyle getHeaderStyle() {
        return headerStyle;
    }


    public int get2007MaxRow() {
        return xlsxMaxRow;
    }

    public int getRowIndex() {
        return bodyRowStartIndex - 1;
    }

    public int getBodyRowCount() {
        return bodyRowCount;
    }

    public List<String> getExcelHeader() {
        return excelHeader;
    }
}
