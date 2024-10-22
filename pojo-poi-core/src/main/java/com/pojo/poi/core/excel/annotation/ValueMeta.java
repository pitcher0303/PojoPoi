package com.pojo.poi.core.excel.annotation;

import com.pojo.poi.core.excel.style.ExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ValueMeta {
    String[] xAxis() default "A";

    int[] yAxis() default 1;

    String value() default "";

    ExcelCellStyle cellStyle() default @ExcelCellStyle;
}
