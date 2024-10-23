package com.pojo.poi.core.excel.style;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface RgbColor {
    int red() default 0;

    int green() default 0;

    int blue() default 0;
}
