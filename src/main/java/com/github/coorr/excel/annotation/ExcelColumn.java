package com.github.coorr.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    String headerName() default "";

    String secondHeaderName() default "";

    String columnName() default "";

    boolean required() default false;

    int row() default 0;
    int column() default 0;
    int width() default 5000;
}
