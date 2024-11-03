package com.pojo.poi.core.excel.style;

import com.pojo.poi.core.excel.ExcelUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelFont {
    boolean bold() default false;

    short[] fontSize() default {};

    IndexedColors[] color() default {IndexedColors.BLACK};

    class Applier {
        public static void apply(Font font, ExcelFont excelFont) {
            if (excelFont == null) return;
            if (excelFont.bold()) font.setBold(true);
            if (ExcelUtils.isApply(excelFont.fontSize())) font.setFontHeightInPoints(excelFont.fontSize()[0]);
            if (ExcelUtils.isApply(excelFont.color())) font.setColor(excelFont.color()[0].index);
        }
    }
}
