package com.okdori;

import com.okdori.style.NoExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

	String headerName() default "";
	boolean mergeCells() default false;
	boolean mergeRows() default false;

	ExcelColumnStyle headerStyle() default @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class);
	ExcelColumnStyle bodyStyle() default @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class);
}
