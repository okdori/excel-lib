package com.okdori.excel;

import com.okdori.ExcelColumn;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
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

    public XSSFWorkbook generateExcel(List<?> dataList) throws IllegalAccessException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        if (dataList.isEmpty()) {
            return workbook;
        }

        int headerRowNum = 0;
        int subHeaderRowNum = 1;
        int dataStartRowNum = 2;

        XSSFRow headerRow = sheet.createRow(headerRowNum);
        XSSFRow subHeaderRow = sheet.createRow(subHeaderRowNum);

        Field[] fields = dataList.get(0).getClass().getDeclaredFields();
        int colIndex = 0;

        XSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        for (Field field : fields) {
            field.setAccessible(true);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

            if (excelColumn != null) {
                XSSFCell headerCell = headerRow.createCell(colIndex);
                headerCell.setCellValue(excelColumn.headerName());
                headerCell.setCellStyle(headerStyle);

                if (excelColumn.mergeCells()) {
                    if (field.getType().isPrimitive() || field.getType().equals(String.class)) {
                        CellRangeAddress verticalMergeRange = new CellRangeAddress(headerRowNum, subHeaderRowNum, colIndex, colIndex);
                        sheet.addMergedRegion(verticalMergeRange);
                        colIndex++;
                    } else {
                        Field[] nestedFields = field.getType().getDeclaredFields();
                        int nestedStartColIndex = colIndex;

                        for (Field nestedField : nestedFields) {
                            nestedField.setAccessible(true);
                            XSSFCell nestedHeaderCell = subHeaderRow.createCell(colIndex);
                            nestedHeaderCell.setCellValue(nestedField.getName());
                            nestedHeaderCell.setCellStyle(headerStyle);
                            colIndex++;
                        }

                        CellRangeAddress horizontalMergeRange = new CellRangeAddress(headerRowNum, headerRowNum, nestedStartColIndex, colIndex - 1);
                        sheet.addMergedRegion(horizontalMergeRange);
                    }
                } else {
                    XSSFCell subHeaderCell = subHeaderRow.createCell(colIndex);
                    subHeaderCell.setCellValue("");
                    subHeaderCell.setCellStyle(headerStyle);
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
                            colIndex++;
                        } else {
                            Field[] nestedFields = value.getClass().getDeclaredFields();
                            int nestedStartColIndex = colIndex;

                            for (Field nestedField : nestedFields) {
                                nestedField.setAccessible(true);
                                XSSFCell nestedDataCell = dataRow.createCell(colIndex);
                                Object nestedValue = nestedField.get(value);
                                nestedDataCell.setCellValue(nestedValue != null ? nestedValue.toString() : "");
                                colIndex++;
                            }
                        }
                    } else {
                        if (value instanceof LocalDate) {
                            dataCell.setCellValue(((LocalDate) value).toString());
                        } else {
                            dataCell.setCellValue(value != null ? value.toString() : "");
                        }

                        colIndex++;
                    }
                }
            }
        }

        for (int i = 0; i < fields.length; i++) {
            sheet.autoSizeColumn(i);
        }

        return workbook;
    }
}
