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
        int dataRowNum = 1;

        XSSFRow headerRow = sheet.createRow(headerRowNum);
        XSSFRow dataRow = sheet.createRow(dataRowNum);

        Field[] fields = dataList.get(0).getClass().getDeclaredFields();
        int colIndex = 0;

        for (Field field : fields) {
            field.setAccessible(true);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

            if (excelColumn != null) {
                XSSFCell cell = headerRow.createCell(colIndex);
                cell.setCellValue(excelColumn.headerName());

                Object value = field.get(dataList.get(0));
                XSSFCell dataCell = dataRow.createCell(colIndex);
                dataCell.setCellValue(value != null ? value.toString() : "");

                if (excelColumn.mergeCells()) {
                    CellRangeAddress mergeRange = new CellRangeAddress(
                            headerRowNum, headerRowNum, colIndex, colIndex + 1
                    );
                    sheet.addMergedRegion(mergeRange);
                    colIndex += 2;
                } else {
                    colIndex++;
                }
            }
        }

        return workbook;
    }
}
