package com.okdori.excel;

import com.okdori.ExcelColumn;
import com.okdori.resource.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.List;

/**
 * packageName    : com.okdori.excel
 * fileName       : ExcelGenerator
 * author         : okdori
 * date           : 2024. 8. 9.
 * description    :
 */

public class ExcelGenerator {
    private String sheetName = "Sheet1";

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public XSSFWorkbook generateExcel(List<?> dataList, Class<?> clazz) throws IllegalAccessException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        ExcelRenderResource resource = ExcelRenderResourceFactory.prepareRenderResource(clazz, workbook, new DefaultDataFormatDecider());

        if (dataList.isEmpty()) {
            return workbook;
        }

//        XSSFRow row = sheet.createRow(0);
//        row.setHeight((short) (17 * 20));
//        sheet.setColumnWidth(0, 2 * 256);

        int headerRowNum = 1;
        int subHeaderRowNum = 2;
        int dataStartRowNum = 3;

        XSSFRow headerRow = sheet.createRow(headerRowNum);
        XSSFRow subHeaderRow = sheet.createRow(subHeaderRowNum);

        Field[] fields = dataList.get(0).getClass().getDeclaredFields();
        int colIndex = 0;

        for (Field field : fields) {
            field.setAccessible(true);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

            if (excelColumn != null) {
                XSSFCell headerCell = headerRow.createCell(colIndex);
                headerCell.setCellValue(excelColumn.headerName());
                headerCell.setCellStyle(resource.getCellStyle(field.getName(),  ExcelRenderLocation.HEADER));

                if (excelColumn.mergeCells()) {
                    if (field.getType().isPrimitive() || field.getType().equals(String.class)
                        || java.time.temporal.Temporal.class.isAssignableFrom(field.getType())
                    ) {
                        CellRangeAddress verticalMergeRange = new CellRangeAddress(headerRowNum, subHeaderRowNum, colIndex, colIndex);
                        sheet.addMergedRegion(verticalMergeRange);
                        colIndex++;

                        XSSFCell subHeaderCell = subHeaderRow.createCell(colIndex - 1);
                        subHeaderCell.setCellStyle(resource.getCellStyle(field.getName(),  ExcelRenderLocation.HEADER));
                    } else {
                        Field[] nestedFields = field.getType().getDeclaredFields();
                        int nestedStartColIndex = colIndex;

                        for (Field nestedField : nestedFields) {
                            nestedField.setAccessible(true);
                            ExcelColumn nestedColumn = nestedField.getAnnotation(ExcelColumn.class);
                            XSSFCell nestedHeaderCell = subHeaderRow.createCell(colIndex);
                            nestedHeaderCell.setCellValue(nestedColumn.headerName());
                            nestedHeaderCell.setCellStyle(resource.getCellStyle(field.getName(),  ExcelRenderLocation.HEADER));
                            colIndex++;

                            XSSFCell horizonHeaderCell = headerRow.createCell(colIndex - 1);
                            horizonHeaderCell.setCellValue(excelColumn.headerName());
                            horizonHeaderCell.setCellStyle(resource.getCellStyle(field.getName(), ExcelRenderLocation.HEADER));
                        }

                        CellRangeAddress horizontalMergeRange = new CellRangeAddress(headerRowNum, headerRowNum, nestedStartColIndex, colIndex - 1);
                        sheet.addMergedRegion(horizontalMergeRange);
                    }
                } else {
                    XSSFCell subHeaderCell = subHeaderRow.createCell(colIndex);
                    subHeaderCell.setCellValue("");
                    subHeaderCell.setCellStyle(resource.getCellStyle(field.getName(),  ExcelRenderLocation.HEADER));
                    colIndex++;
                }
            }
        }

        for (int rowNum = 0; rowNum < dataList.size(); rowNum++) {
            XSSFRow dataRow = sheet.createRow(dataStartRowNum + rowNum);
            Object dataObject = dataList.get(rowNum);

            colIndex = 0;

            for (Field field : fields) {
                field.setAccessible(true);
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

                if (excelColumn != null) {
                    XSSFCell dataCell = dataRow.createCell(colIndex);
                    Object value = field.get(dataObject);

                    if (excelColumn.mergeCells()) {
                        if (field.getType().isPrimitive() || field.getType().equals(String.class)) {
                            dataCell.setCellValue(value != null ? value.toString() : "");
                            dataCell.setCellStyle(resource.getCellStyle(field.getName(), ExcelRenderLocation.BODY));
                            colIndex = autoSizeColumn(sheet, colIndex);
                        } else {
                            Field[] nestedFields = value.getClass().getDeclaredFields();
                            for (Field nestedField : nestedFields) {
                                nestedField.setAccessible(true);
                                XSSFCell nestedDataCell = dataRow.createCell(colIndex);
                                Object nestedValue = nestedField.get(value);
                                nestedDataCell.setCellValue(nestedValue != null ? nestedValue.toString() : "");
                                nestedDataCell.setCellStyle(resource.getCellStyle(field.getName(), ExcelRenderLocation.BODY));
                                colIndex = autoSizeColumn(sheet, colIndex);
                            }
                        }
                    } else {
                        if (value instanceof LocalDate) {
                            dataCell.setCellValue(((LocalDate) value).toString());
                            dataCell.setCellStyle(resource.getCellStyle(field.getName(), ExcelRenderLocation.BODY));
                        } else {
                            dataCell.setCellValue(value != null ? value.toString() : "");
                            dataCell.setCellStyle(resource.getCellStyle(field.getName(), ExcelRenderLocation.BODY));
                        }

                        colIndex = autoSizeColumn(sheet, colIndex);
                    }
                }
            }
        }

        return workbook;
    }

    private int autoSizeColumn(XSSFSheet sheet, int colIndex) {
        sheet.autoSizeColumn(colIndex);
        sheet.setColumnWidth(colIndex, (sheet.getColumnWidth(colIndex)) + 1024);
        return colIndex + 1;
    }
}
