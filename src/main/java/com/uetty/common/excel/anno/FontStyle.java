package com.uetty.common.excel.anno;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface FontStyle {

    String name() default "";

    double size() default -1;

    IndexedColors color() default IndexedColors.AUTOMATIC;

    boolean bold() default false;
}