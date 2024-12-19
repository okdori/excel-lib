package com.okdori.excel;

import com.okdori.ExcelColumn;
import com.okdori.resource.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.*;
import java.io.IOException;

/**
 * packageName    : com.okdori.excel
 * fileName       : NewExcelGenerator
 * author         : okdori
 * date           : 2024. 12. 19.
 * description    :
 */

public class NewExcelGenerator {
    private String sheetName = "Sheet1";
    private static final int FLUSH_THRESHOLD = 1000;
    private static final int WINDOW_SIZE = 1000; // SXSSFWorkbook memory window size

    private static class FieldInfo {
        Field field;
        ExcelColumn annotation;
        List<FieldInfo> nestedFields;
        boolean isPrimitiveOrSimple;

        FieldInfo(Field field, ExcelColumn annotation) {
            this.field = field;
            this.annotation = annotation;
            this.isPrimitiveOrSimple = isPrimitiveOrSimpleType(field);
            this.nestedFields = new ArrayList<>();
        }
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    private static boolean isPrimitiveOrSimpleType(Field field) {
        return field.getType().isPrimitive()
                || field.getType().equals(String.class)
                || java.time.temporal.Temporal.class.isAssignableFrom(field.getType())
                || Number.class.isAssignableFrom(field.getType());
    }

    private List<FieldInfo> analyzeClass(Class<?> clazz) {
        List<FieldInfo> fieldInfos = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();

        for (Field field : fields) {
            field.setAccessible(true);
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);

            if (annotation != null) {
                FieldInfo fieldInfo = new FieldInfo(field, annotation);

                if (annotation.mergeCells() && !fieldInfo.isPrimitiveOrSimple) {
                    Field[] nestedFields = field.getType().getDeclaredFields();
                    for (Field nestedField : nestedFields) {
                        nestedField.setAccessible(true);
                        ExcelColumn nestedAnnotation = nestedField.getAnnotation(ExcelColumn.class);
                        if (nestedAnnotation != null) {
                            fieldInfo.nestedFields.add(new FieldInfo(nestedField, nestedAnnotation));
                        }
                    }
                }
                fieldInfos.add(fieldInfo);
            }
        }
        return fieldInfos;
    }

    public Workbook generateExcel(List<?> dataList, Class<?> clazz) throws IllegalAccessException {
        if (dataList.isEmpty()) {
            return new SXSSFWorkbook(WINDOW_SIZE);
        }

        // SXSSFWorkbook create (memory optimizer)
        SXSSFWorkbook workbook = new SXSSFWorkbook(WINDOW_SIZE);
        try {
            workbook.setCompressTempFiles(true); // compress temp files
        } catch (Exception e) {
            // pass
        }

        Sheet sheet = workbook.createSheet(sheetName);

        // ready style resource
        ExcelRenderResource resource = ExcelRenderResourceFactory.prepareRenderResource(clazz, workbook, new DefaultDataFormatDecider());

        // cache fieldInfos
        List<FieldInfo> fieldInfos = analyzeClass(clazz);

        // create header
        createHeaders(sheet, fieldInfos, resource);

        // create data row
        processDataRows(sheet, dataList, fieldInfos, resource);

        optimizeColumnWidths(sheet, fieldInfos);

        return workbook;
    }

    private void createHeaders(Sheet sheet, List<FieldInfo> fieldInfos, ExcelRenderResource resource) {
        Row headerRow = sheet.createRow(0);
        Row subHeaderRow = sheet.createRow(1);
        int colIndex = 0;

        for (FieldInfo fieldInfo : fieldInfos) {
            if (fieldInfo.annotation.mergeCells()) {
                if (fieldInfo.isPrimitiveOrSimple) {
                    createSimpleHeaderCell(sheet, headerRow, subHeaderRow, colIndex, fieldInfo, resource);
                    colIndex++;
                } else {
                    colIndex = createNestedHeaderCells(sheet, headerRow, subHeaderRow, colIndex, fieldInfo, resource);
                }
            } else {
                createSimpleHeaderCell(sheet, headerRow, subHeaderRow, colIndex, fieldInfo, resource);
                colIndex++;
            }
        }
    }

    private void createSimpleHeaderCell(Sheet sheet, Row headerRow, Row subHeaderRow,
                                        int colIndex, FieldInfo fieldInfo, ExcelRenderResource resource) {
        Cell headerCell = headerRow.createCell(colIndex);
        headerCell.setCellValue(fieldInfo.annotation.headerName());
        headerCell.setCellStyle(resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER));

        Cell subHeaderCell = subHeaderRow.createCell(colIndex);
        subHeaderCell.setCellStyle(resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER));

        if (fieldInfo.annotation.mergeCells()) {
            sheet.addMergedRegion(new CellRangeAddress(0, 1, colIndex, colIndex));
        }
    }

    private int createNestedHeaderCells(Sheet sheet, Row headerRow, Row subHeaderRow,
                                        int colIndex, FieldInfo fieldInfo, ExcelRenderResource resource) {
        int startColIndex = colIndex;

        for (FieldInfo nestedField : fieldInfo.nestedFields) {
            Cell subHeaderCell = subHeaderRow.createCell(colIndex);
            subHeaderCell.setCellValue(nestedField.annotation.headerName());
            subHeaderCell.setCellStyle(resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER));

            Cell headerCell = headerRow.createCell(colIndex);
            headerCell.setCellValue(fieldInfo.annotation.headerName());
            headerCell.setCellStyle(resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER));

            colIndex++;
        }

        if (colIndex > startColIndex) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIndex, colIndex - 1));
        }

        return colIndex;
    }

    private void processDataRows(Sheet sheet, List<?> dataList, List<FieldInfo> fieldInfos,
                                 ExcelRenderResource resource) throws IllegalAccessException {
        int rowCount = 0;

        for (Object dataObject : dataList) {
            Row dataRow = sheet.createRow(rowCount + 2);
            int colIndex = 0;

            for (FieldInfo fieldInfo : fieldInfos) {
                if (fieldInfo.annotation.mergeCells()) {
                    if (fieldInfo.isPrimitiveOrSimple) {
                        createSimpleDataCell(dataRow, colIndex, fieldInfo, dataObject, resource);
                        colIndex++;
                    } else {
                        colIndex = createNestedDataCells(dataRow, colIndex, fieldInfo, dataObject, resource);
                    }
                } else {
                    createSimpleDataCell(dataRow, colIndex, fieldInfo, dataObject, resource);
                    colIndex++;
                }
            }

            rowCount++;
        }
    }

    private void createSimpleDataCell(Row dataRow, int colIndex, FieldInfo fieldInfo,
                                      Object dataObject, ExcelRenderResource resource) throws IllegalAccessException {
        Cell cell = dataRow.createCell(colIndex);
        Object value = fieldInfo.field.get(dataObject);

        if (value instanceof LocalDate) {
            cell.setCellValue(value.toString());
        } else {
            cell.setCellValue(value != null ? value.toString() : "");
        }

        cell.setCellStyle(resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.BODY));
    }

    private int createNestedDataCells(Row dataRow, int colIndex, FieldInfo fieldInfo,
                                      Object dataObject, ExcelRenderResource resource) throws IllegalAccessException {
        Object nestedObject = fieldInfo.field.get(dataObject);

        for (FieldInfo nestedField : fieldInfo.nestedFields) {
            Cell cell = dataRow.createCell(colIndex);

            if (nestedObject != null) {
                Object value = nestedField.field.get(nestedObject);
                cell.setCellValue(value != null ? value.toString() : "");
            } else {
                cell.setCellValue("");
            }

            cell.setCellStyle(resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.BODY));
            colIndex++;
        }

        return colIndex;
    }

    private void optimizeColumnWidths(Sheet sheet, List<FieldInfo> fieldInfos) {
        int totalColumns = getTotalColumnCount(fieldInfos);
        for (int i = 0; i < totalColumns; i++) {
            sheet.setColumnWidth(i, 256 * 15);
        }
    }

    private int getTotalColumnCount(List<FieldInfo> fieldInfos) {
        int count = 0;
        for (FieldInfo fieldInfo : fieldInfos) {
            if (fieldInfo.annotation.mergeCells() && !fieldInfo.isPrimitiveOrSimple) {
                count += fieldInfo.nestedFields.size();
            } else {
                count++;
            }
        }
        return count;
    }

    public void dispose(Workbook workbook) {
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
        try {
            workbook.close();
        } catch (IOException e) {
            // pass
        }
    }
}
