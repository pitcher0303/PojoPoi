package com.pojo.poi.core.excel.annotation;

import com.pojo.poi.core.excel.style.ExcelRowStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface RowMeta {
    RowType rowType() default RowType.Y_RANDOM;

    String[] xAxis() default "A";

    int[] yAxis() default 1;

    int startYAxis() default 1;

    int endYAxis() default 1;

    Class<?> target();

    ValueMeta[] headerMetas() default {};

    GroupByMeta[] groupBys() default {};

    ExcelRowStyle[] rowStyle() default {};

    MetaOrder metaOrder() default @MetaOrder;

    enum RowType {
        Y_RANDOM,
        //TODO : X_RANDOM, RANGE 추가 예정.
//        X_RANDOM,
//        RANGE,
    }
}
