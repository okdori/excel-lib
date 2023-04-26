package com.okdori.excel.onesheet;

import com.okdori.excel.SXSSFExcelFile;
import com.okdori.resource.DataFormatDecider;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.List;

/**
 * OneSheetExcelFile
 *
 * - support Excel Version over 2007
 * - support one sheet rendering
 * - support different DataFormat by Class Type
 * - support Custom CellStyle according to (header or body) and data field
 */
public final class OneSheetExcelFile<T> extends SXSSFExcelFile<T> {

	private static final int ROW_START_INDEX = 0;
	private static final int COLUMN_START_INDEX = 0;
	private int currentRowIndex = ROW_START_INDEX;

	public OneSheetExcelFile(Class<T> type) {
		super(type);
	}

	public OneSheetExcelFile(List<T> data, Class<T> type) {
		super(data, type);
	}

	public OneSheetExcelFile(List<T> data, Class<T> type, DataFormatDecider dataFormatDecider) {
		super(data, type, dataFormatDecider);
	}

	@Override
	protected void validateData(List<T> data) {
		int maxRows = supplyExcelVersion.getMaxRows();
		if (data.size() > maxRows) {
			throw new IllegalArgumentException(
					String.format("This concrete ExcelFile does not support over %s rows", maxRows));
		}
	}

	@Override
	public void renderExcel(List<T> data) {
		// 1. Create sheet and renderHeader
		sheet = wb.createSheet();
		((SXSSFSheet)sheet).trackAllColumnsForAutoSizing();
		renderHeadersWithNewSheet(sheet, currentRowIndex++, COLUMN_START_INDEX);

		if (data.isEmpty()) {
			return;
		}

		// 2. Render Body
		for (Object renderedData : data) {
			renderBody(renderedData, currentRowIndex++, COLUMN_START_INDEX);
		}
	}

	@Override
	public void addRows(List<T> data) {
		renderBody(data, currentRowIndex++, COLUMN_START_INDEX);
	}

}
