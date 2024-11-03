package com.pojo.poi.core.excel.annotation;

import com.pojo.poi.core.excel.style.ExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellMeta {
    String[] xAxis() default "A";

    int[] yAxis() default 1;

    ValueMeta headerMeta() default @ValueMeta;

    ExcelCellStyle[] cellStyle() default {};

    MetaOrder metaOrder() default @MetaOrder;
}
