package com.pojo.poi.core.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;

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

    enum Type {
        indexed,
        rgb
    }

    class Applier {
        public static void apply(CellStyle cellStyle, ExcelColor excelColor) {
            if (excelColor == null) return;
            switch (excelColor.type()) {
                case indexed -> {
                    cellStyle.setFillForegroundColor(excelColor.indexedColor().index);
                }
                case rgb -> {
                    RgbColor rgbColor = excelColor.rgbColor();
                    cellStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) rgbColor.red(), (byte) rgbColor.green(), (byte) rgbColor.blue()}));
                }
            }
        }
    }
}
