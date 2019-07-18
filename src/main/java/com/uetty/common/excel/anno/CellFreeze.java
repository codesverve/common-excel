package com.uetty.common.excel.anno;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface CellFreeze {

    int freezeCol() default 0;

    int freezeRow() default 0;
}
