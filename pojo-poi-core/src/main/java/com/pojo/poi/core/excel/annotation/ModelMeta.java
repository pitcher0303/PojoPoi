package com.pojo.poi.core.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ModelMeta {
    public ModelType type() default ModelType.NORMAL;

    String[] xAxis() default "A";

    int[] yAxis() default 1;

    public int startYAxis() default 1;

    public int endYAxis() default 1;

    public ValueMeta[] headerMetas() default {};

    public enum ModelType {NORMAL}
}
