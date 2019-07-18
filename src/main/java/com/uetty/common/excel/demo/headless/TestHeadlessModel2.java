package com.uetty.common.excel.demo.headless;

import com.uetty.common.excel.demo.MyConstraintEnum;
import com.uetty.common.excel.easyexcel.hssf.XlsExcelWriter;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

public class TestHeadlessModel2 {

    // 与TestHeadlessModel1相比，注解不一样，将@ExcelProperty换成了@ExcelColumnNum
    public static void main(String[] args) throws IOException {
        XlsExcelWriter writer = new XlsExcelWriter("/data/test3.xls", MyModel2.class, 0);


        // 数据写入excel
        writer.write(createTestListObject2());
    }

    private static final int ROW_SIZE = 14;
    private static String[] PROP1_VALUES = {"aaa1", "aaa2", "aaa3"};
    private static MyConstraintEnum[] enValues = MyConstraintEnum.values();
    private static List<MyModel2> createTestListObject2() {
        List<MyModel2> list = new ArrayList<>();
        for (int i = 0; i < ROW_SIZE; i++) {
            MyModel2 ts = new MyModel2();
            ts.setPr1("pr1 == " + i);
            int index = i + (Math.random() > 0.8 ? 1 : 0);
            ts.setPropValue1(PROP1_VALUES[index % PROP1_VALUES.length]);
            ts.setScore((int) (Math.random() * 100));
            ts.setPr2(UUID.randomUUID().toString());
            ts.setPropValue2(enValues[index % enValues.length].getValue());
            ts.setDate(new Date());
            list.add(ts);
        }
        return list;
    }
}
