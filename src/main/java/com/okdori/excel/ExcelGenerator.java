package com.okdori.excel;

import com.okdori.ExcelColumn;
import com.okdori.resource.*;
import com.okdori.utils.TypeUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.IOException;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.*;

/**
 * packageName    : com.okdori.excel
 * fileName       : ExcelGenerator
 * author         : okdori
 * date           : 2024. 12. 20.
 * description    :
 */


@Getter
@Setter
public class ExcelGenerator {
    private static final int FLUSH_THRESHOLD = 1000;
    private static final int WINDOW_SIZE = 1000;
    private static final int DEFAULT_HEIGHT = 17;
    private String sheetName = "Sheet1";
    private SXSSFWorkbook workbook;

    @Getter
    private static class FieldInfo {
        private final Field field;
        private final ExcelColumn annotation;
        private final List<FieldInfo> nestedFields;
        private final boolean isPrimitiveOrSimple;

        private FieldInfo(Field field, ExcelColumn annotation) {
            this.field = field;
            this.annotation = annotation;
            this.isPrimitiveOrSimple = TypeUtils.isPrimitiveOrSimpleType(field);
            this.nestedFields = new ArrayList<>();
        }

        public static FieldInfo create(Field field, ExcelColumn annotation) {
            return new FieldInfo(field, annotation);
        }
    }

    public Workbook generateExcel(List<?> dataList, Class<?> clazz) throws IllegalAccessException {
        initializeWorkbook(dataList);
        if (dataList.isEmpty()) {
            return this.workbook;
        }

        Sheet sheet = createAndConfigureSheet();
        ExcelRenderResource resource = prepareRenderResource(clazz);
        List<FieldInfo> fieldInfos = analyzeClass(clazz);

        processExcelGeneration(sheet, dataList, fieldInfos, resource);

        return this.workbook;
    }

    private void initializeWorkbook(List<?> dataList) {
        this.workbook = new SXSSFWorkbook(WINDOW_SIZE);
        if (!dataList.isEmpty()) {
            configureWorkbook();
        }
    }

    private void configureWorkbook() {
        try {
            this.workbook.setCompressTempFiles(true);
        } catch (Exception ignored) {
            // Compression configuration failed - continuing with default settings
        }
    }

    private Sheet createAndConfigureSheet() {
        return workbook.createSheet(sheetName);
    }

    private ExcelRenderResource prepareRenderResource(Class<?> clazz) {
        return ExcelRenderResourceFactory.prepareRenderResource(
                clazz,
                this.workbook,
                new DefaultDataFormatDecider()
        );
    }

    private List<FieldInfo> analyzeClass(Class<?> clazz) {
        List<FieldInfo> fieldInfos = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();

        Arrays.stream(fields)
                .filter(this::isValidExcelField)
                .map(this::createFieldInfo)
                .forEach(fieldInfo -> {
                    processNestedFields(fieldInfo);
                    fieldInfos.add(fieldInfo);
                });

        return fieldInfos;
    }

    private boolean isValidExcelField(Field field) {
        field.setAccessible(true);
        return field.getAnnotation(ExcelColumn.class) != null;
    }

    private FieldInfo createFieldInfo(Field field) {
        ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
        return FieldInfo.create(field, annotation);
    }

    private void processNestedFields(FieldInfo fieldInfo) {
        if (!shouldProcessNestedFields(fieldInfo)) {
            return;
        }

        Arrays.stream(fieldInfo.getField().getType().getDeclaredFields())
                .filter(this::isValidExcelField)
                .map(this::createFieldInfo)
                .forEach(fieldInfo.getNestedFields()::add);
    }

    private boolean shouldProcessNestedFields(FieldInfo fieldInfo) {
        return fieldInfo.getAnnotation().mergeCells() && !fieldInfo.isPrimitiveOrSimple();
    }

    private void processExcelGeneration(Sheet sheet, List<?> dataList,
                                        List<FieldInfo> fieldInfos,
                                        ExcelRenderResource resource) throws IllegalAccessException {
        createHeaders(sheet, fieldInfos, resource);
        processDataRows(sheet, dataList, fieldInfos, resource);
        optimizeColumnWidths(sheet, fieldInfos);
    }

    private void createHeaders(Sheet sheet, List<FieldInfo> fieldInfos, ExcelRenderResource resource) {
        Row headerRow = sheet.createRow(0);
        Row subHeaderRow = sheet.createRow(1);

        headerRow.setHeight((short)(DEFAULT_HEIGHT * 20));
        subHeaderRow.setHeight((short)(DEFAULT_HEIGHT * 20));

        int colIndex = 0;
        int maxHeaderLines = 1;
        int maxSubHeaderLines = 1;

        for (FieldInfo fieldInfo : fieldInfos) {
            if (fieldInfo.annotation.mergeCells()) {
                if (fieldInfo.isPrimitiveOrSimple) {
                    String headerText = fieldInfo.annotation.headerName();
                    maxHeaderLines = Math.max(maxHeaderLines, headerText.split("\n").length);
                } else {
                    for (FieldInfo nestedField : fieldInfo.nestedFields) {
                        String subHeaderText = nestedField.annotation.headerName();
                        maxSubHeaderLines = Math.max(maxSubHeaderLines, subHeaderText.split("\n").length);
                    }
                    String headerText = fieldInfo.annotation.headerName();
                    maxHeaderLines = Math.max(maxHeaderLines, headerText.split("\n").length);
                }
            } else {
                String headerText = fieldInfo.annotation.headerName();
                maxHeaderLines = Math.max(maxHeaderLines, headerText.split("\n").length);
            }
        }

        headerRow.setHeight((short)(DEFAULT_HEIGHT * 20 * maxHeaderLines));
        subHeaderRow.setHeight((short)(DEFAULT_HEIGHT * 20 * maxSubHeaderLines));

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

        String headerText = fieldInfo.annotation.headerName();
        headerCell.setCellValue(createRichTextString(headerText));

        CellStyle headerStyle = resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER);
        headerStyle.setWrapText(true);
        headerCell.setCellStyle(headerStyle);

        Cell subHeaderCell = subHeaderRow.createCell(colIndex);
        CellStyle subHeaderStyle = resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER);
        subHeaderStyle.setWrapText(true);
        subHeaderCell.setCellStyle(subHeaderStyle);

        if (fieldInfo.annotation.mergeCells()) {
            sheet.addMergedRegion(new CellRangeAddress(0, 1, colIndex, colIndex));
        }
    }

    private int createNestedHeaderCells(Sheet sheet, Row headerRow, Row subHeaderRow,
                                        int colIndex, FieldInfo fieldInfo, ExcelRenderResource resource) {
        int startColIndex = colIndex;

        for (FieldInfo nestedField : fieldInfo.nestedFields) {
            Cell subHeaderCell = subHeaderRow.createCell(colIndex);
            String subHeaderText = nestedField.annotation.headerName();
            subHeaderCell.setCellValue(createRichTextString(subHeaderText));

            CellStyle subHeaderStyle = resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER);
            subHeaderStyle.setWrapText(true);
            subHeaderCell.setCellStyle(subHeaderStyle);

            Cell headerCell = headerRow.createCell(colIndex);
            String headerText = fieldInfo.annotation.headerName();
            headerCell.setCellValue(createRichTextString(headerText));

            CellStyle headerStyle = resource.getCellStyle(fieldInfo.field.getName(), ExcelRenderLocation.HEADER);
            headerStyle.setWrapText(true);
            headerCell.setCellStyle(headerStyle);

            colIndex++;
        }

        if (colIndex > startColIndex) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, startColIndex, colIndex - 1));
        }

        return colIndex;
    }

    private RichTextString createRichTextString(String text) {
        if (this.workbook == null) {
            throw new IllegalStateException("Workbook is not initialized. Please call generateExcel first.");
        }

        return new XSSFRichTextString(text);
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
        SXSSFSheet sxssfSheet = (SXSSFSheet) sheet;

        Row headerRow = sheet.getRow(0);
        Row subHeaderRow = sheet.getRow(1);
        boolean hasSubHeader = hasSubHeaders(fieldInfos);

        int totalRows = sheet.getLastRowNum();
        int sampleSize = 1000;
        int samplingInterval = Math.max(1, totalRows / sampleSize);

        for (int i = 0; i < totalColumns; i++) {
            try {
                int maxWidth = 0;

                if (headerRow != null && headerRow.getCell(i) != null) {
                    maxWidth = getContentWidth(headerRow.getCell(i).toString());
                }
                if (hasSubHeader && subHeaderRow != null && subHeaderRow.getCell(i) != null) {
                    maxWidth = Math.max(maxWidth, getContentWidth(subHeaderRow.getCell(i).toString()));
                }

                int startRow = hasSubHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= totalRows; rowNum += samplingInterval) {
                    Row row = sheet.getRow(rowNum);
                    if (row != null && row.getCell(i) != null) {
                        int width = getContentWidth(row.getCell(i).toString());
                        maxWidth = Math.max(maxWidth, width);
                    }
                }

                maxWidth = Math.max(8, Math.min(50, maxWidth));
                sheet.setColumnWidth(i, 256 * maxWidth);

                if (i % 10 == 0) {
                    sxssfSheet.flushRows();
                }
            } catch (Exception e) {
                sheet.setColumnWidth(i, 256 * 15);
            }
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

    private int getContentWidth(String content) {
        int width = 0;
        for (char c : content.toCharArray()) {
            if (Character.UnicodeBlock.of(c) == Character.UnicodeBlock.HANGUL_SYLLABLES) {
                width += 2;
            } else {
                width += 1;
            }
        }
        return width;
    }

    private boolean hasSubHeaders(List<FieldInfo> fieldInfos) {
        return fieldInfos.stream()
                .anyMatch(fieldInfo -> fieldInfo.getAnnotation().mergeCells()
                        && !fieldInfo.isPrimitiveOrSimple());
    }

    public void dispose() {
        if (this.workbook != null) {
            this.workbook.dispose();
            try {
                this.workbook.close();
            } catch (IOException e) {
                // pass
            }
        }
    }
}
