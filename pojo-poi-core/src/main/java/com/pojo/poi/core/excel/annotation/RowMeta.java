package com.pojo.poi.core.excel.annotation;

import com.pojo.poi.core.excel.style.ExcelRowStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface RowMeta {
    public RowType rowType() default RowType.RANGE;

    String[] xAxis() default "A";

    int[] yAxis() default 1;

    public int startYAxis() default 1;

    public int endYAxis() default 1;

    public Class<?> target();

    public ValueMeta[] headerMetas() default {};

    public GroupByMeta[] groupBys() default {};

    public ExcelRowStyle[] rowStyle() default {};

    public enum RowType {RANGE, Y_RANDOM, X_RANDOM}
}
