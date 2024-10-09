package org.example.excel.annotation;

import org.example.excel.model.ExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface GroupByMeta {
    Type type() default Type.NONE;

    DataType dataType() default DataType.AUTO_INCREMENT;

    String[] xAxis() default "A";

    int[] yAxis() default 1;

    ExcelCellStyle cellStyle() default ExcelCellStyle.NONE;

    public enum Type {NONE, Y_MERGE}

    public enum DataType {AUTO_INCREMENT, CELL_DATA}
}