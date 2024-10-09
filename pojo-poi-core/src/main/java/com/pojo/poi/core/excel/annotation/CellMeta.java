package org.example.excel.annotation;

import org.example.excel.model.ExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellMeta {
    AxisType axisType() default AxisType.NORMAL;

    String[] xAxis() default "A";

    int[] yAxis() default 1;

    ValueMeta headerMeta() default @ValueMeta;

    ExcelCellStyle cellStyle() default ExcelCellStyle.NONE;

    public enum AxisType {NORMAL, Y_RANDOM, X_RANDOM}
}
