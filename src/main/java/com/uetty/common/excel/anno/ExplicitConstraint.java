package com.uetty.common.excel.anno;

import com.uetty.common.excel.constant.ConstraintValue;
import com.uetty.common.excel.constant.NoneConstraint;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExplicitConstraint {

    String[] source() default {};

    Class<? extends Enum<? extends ConstraintValue>> enumSource() default NoneConstraint.class;
}
