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
 * Row 테두리만 현재 허용.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelRowStyle {
    BorderStyle[] borderTop() default {};

    BorderStyle[] borderRight() default {};

    BorderStyle[] borderBottom() default {};

    BorderStyle[] borderLeft() default {};
}
