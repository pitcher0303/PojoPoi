package com.pojo.poi.core.excel.style;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellFont {
    boolean bold() default false;
    ExcelColor color() default @ExcelColor(indexedColor = IndexedColors.BLACK);
}
