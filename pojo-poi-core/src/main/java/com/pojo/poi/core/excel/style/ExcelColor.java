package com.pojo.poi.core.excel.style;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColor {
    Type type() default Type.indexed;
    IndexedColors indexedColor() default IndexedColors.WHITE;
    RgbColor rgbColor() default @RgbColor();

    public enum Type {
        indexed,
        rgb
    }
}
