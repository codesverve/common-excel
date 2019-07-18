package com.uetty.common.excel.demo;

import com.uetty.common.excel.constant.ConstraintValue;

public enum MyConstraintEnum implements ConstraintValue {

    PROP2_A("prop_a"),
    PROP2_B("prop_b"),
    PROP2_C("prop_c"),
    PROP2_D("prop_d"),
    ;

    String value;

    MyConstraintEnum(String value) {
        this.value = value;
    }

    @Override
    public String getValue() {
        return value;
    }
}
