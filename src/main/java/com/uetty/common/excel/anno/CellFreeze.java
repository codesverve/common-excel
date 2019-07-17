package com.uetty.common.excel.anno;

import java.lang.annotation.*;

/**
 * @Author: Vince
 * @Date: 2019/7/17 20:51
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface CellFreeze {

    int freezeCol() default 0;

    int freezeRow() default 0;
}
