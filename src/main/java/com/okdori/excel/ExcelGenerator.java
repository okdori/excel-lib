package com.okdori.excel;

import com.okdori.ExcelColumn;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.util.List;

/**
 * packageName    : com.okdori.excel
 * fileName       : ExcelGenerator
 * author         : okdori
 * date           : 2024. 8. 9.
 * description    :
 */

public class ExcelGenerator {
    public XSSFWorkbook generateExcel(String sheetName, List<?> dataList) throws IllegalAccessException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(sheetName);

        if (dataList.isEmpty()) {
            return workbook;
        }

        int headerRowNum = 0;
        int subHeaderRowNum = 1;
        int dataRowNum = 2;

        XSSFRow headerRow = sheet.createRow(headerRowNum);
        XSSFRow subHeaderRow = sheet.createRow(subHeaderRowNum);
        XSSFRow dataRow = sheet.createRow(dataRowNum);

        Field[] fields = dataList.get(0).getClass().getDeclaredFields();
        int colIndex = 0;

        for (Field field : fields) {
            field.setAccessible(true);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

            if (excelColumn != null) {
                Object value = field.get(dataList.get(0));

                if (excelColumn.isNestedObject() && value != null) {
                    XSSFCell mergeCell = headerRow.createCell(colIndex);
                    mergeCell.setCellValue(excelColumn.headerName());

                    Field[] nestedFields = field.getType().getDeclaredFields();
                    int nestedStartColIndex = colIndex;

                    for (Field nestedField : nestedFields) {
                        nestedField.setAccessible(true);
                        XSSFCell nestedHeaderCell = subHeaderRow.createCell(colIndex);
                        nestedHeaderCell.setCellValue(nestedField.getName());

                        Object nestedValue = nestedField.get(value);
                        XSSFCell nestedDataCell = dataRow.createCell(colIndex);
                        nestedDataCell.setCellValue(nestedValue != null ? nestedValue.toString() : "");

                        colIndex++;
                    }

                    CellRangeAddress mergeRange = new CellRangeAddress(headerRowNum, headerRowNum, nestedStartColIndex, colIndex - 1);
                    sheet.addMergedRegion(mergeRange);

                } else {
                    XSSFCell cell = headerRow.createCell(colIndex);
                    cell.setCellValue(excelColumn.headerName());

                    XSSFCell subHeaderCell = subHeaderRow.createCell(colIndex);
                    subHeaderCell.setCellValue("");

                    XSSFCell dataCell = dataRow.createCell(colIndex);
                    dataCell.setCellValue(value != null ? value.toString() : "");

                    CellRangeAddress verticalMergeRange = new CellRangeAddress(headerRowNum, subHeaderRowNum, colIndex, colIndex);
                    sheet.addMergedRegion(verticalMergeRange);

                    colIndex++;
                }
            }
        }

        return workbook;
    }
}
