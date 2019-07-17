package com.uetty.common.excel.anno;

import java.lang.annotation.*;

/**
 * @Author: Vince
 * @Date: 2019/7/11 16:03
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface ColumnWidth {

    int width() default -1;

}
