package com.uetty.common.excel.anno;

import java.lang.annotation.*;

// 暂时没用，计划自己实现POI的封装
@Deprecated
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {
    String[] value() default {""};

    int index() default 0;

    String format() default "";
}
