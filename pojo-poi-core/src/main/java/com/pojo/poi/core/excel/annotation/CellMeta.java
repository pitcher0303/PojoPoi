package com.pojo.poi.core.excel.annotation;

import com.pojo.poi.core.excel.model.ExcelCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.util.function.Consumer;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellMeta {
    String[] xAxis() default "A";

    int[] yAxis() default 1;

    ValueMeta headerMeta() default @ValueMeta;

    ExcelCellStyle cellStyle() default ExcelCellStyle.NONE;

    boolean fitSize() default false;
}
