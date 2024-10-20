package com.pojo.poi.core.excel.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public @interface CellStyle {
    boolean wrapText() default false;
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.LEFT;
    VerticalAlignment verticalAlignment() default VerticalAlignment.CENTER;
    BorderStyle borderTop() default BorderStyle.NONE;
    BorderStyle borderRight() default BorderStyle.NONE;
    BorderStyle borderBottom() default BorderStyle.NONE;
    BorderStyle borderLeft() default BorderStyle.NONE;
    ColorRGB backgroundColor() default @ColorRGB;
    CellFont font() default @CellFont;
}
