package com.pojo.poi.core.excel.annotation;

import com.pojo.poi.core.excel.style.ExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface GroupByMeta {
    DataType dataType() default DataType.AUTO_INCREMENT;

    String[] xAxis() default "A";

    int[] yAxis() default 1;

    ExcelCellStyle[] cellStyle() default {};

    public enum DataType {AUTO_INCREMENT, CELL_DATA}
}
