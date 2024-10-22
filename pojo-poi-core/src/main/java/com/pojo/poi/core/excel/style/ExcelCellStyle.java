package com.pojo.poi.core.excel.style;

import org.apache.poi.ss.usermodel.*;

import java.lang.annotation.*;

/**
 * 기본적인 셀 스타일
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellStyle {
    boolean wrapText() default false;
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.LEFT;
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;
    BorderStyle borderTop() default BorderStyle.NONE;
    BorderStyle borderRight() default BorderStyle.NONE;
    BorderStyle borderBottom() default BorderStyle.NONE;
    BorderStyle borderLeft() default BorderStyle.NONE;
    FillPatternType fillPattern() default FillPatternType.SOLID_FOREGROUND;
    ExcelCellFont font() default @ExcelCellFont(color = @ExcelColor(indexedColor = IndexedColors.BLACK));
    ExcelColor foregroundColor() default @ExcelColor(indexedColor = IndexedColors.WHITE1);
    ExcelColor backgroundColor() default @ExcelColor(indexedColor = IndexedColors.WHITE1);
}
