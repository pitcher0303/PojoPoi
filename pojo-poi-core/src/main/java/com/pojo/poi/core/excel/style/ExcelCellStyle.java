package com.pojo.poi.core.excel.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 기본적인 셀 스타일
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellStyle {
    boolean wrapText() default false;

    HorizontalAlignment[] horizontalAlignment() default {};

    VerticalAlignment[] verticalAlignment() default {VerticalAlignment.CENTER};

    BorderStyle[] borderTop() default {};

    BorderStyle[] borderRight() default {};

    BorderStyle[] borderBottom() default {};

    BorderStyle[] borderLeft() default {};

    FillPatternType[] fillPattern() default {};

    ExcelColor[] foregroundColor() default {};

    ExcelColor[] backgroundColor() default {};

    ExcelFont[] font() default {};
}
