package com.pojo.poi.core.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface MetaOrder {
    Type type() default Type.NONE;
    int value() default Integer.MAX_VALUE;
    int referenceMetaOrder() default Integer.MAX_VALUE;

    enum Type {
        NONE,
        //지원 예정. referenceMetaOrder 마지막 XAxis 값을 참조함
//        X_REFERENCES,
        Y_REFERENCES,
        //지원 예정. referenceMetaOrder 마지막 XAxis, YAxis 값을 참조함
//        XY_REFERENCES,
    }
}
