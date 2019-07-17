package com.uetty.common.excel.demo.headless;

import com.uetty.common.excel.demo.MyConstraintEnum;
import com.uetty.common.excel.demo.MyModel1;
import com.uetty.common.excel.easyexcel.hssf.XlsExcelWriter;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

/**
 * @Author: Vince
 * @Date: 2019/7/17 21:06
 */
public class TestHeadlessModel1 {

    public static void main(String[] args) throws IOException {
        XlsExcelWriter writer = new XlsExcelWriter("/data/test2.xls", MyModel1.class, 0);

        // 设置无需标题行
        writer.setNeedHead(false)
                .addExplicitConstraint(3, 4, 2, 4, new String[]{"理学", "文学", "工学"})
                .addCustomCellStyleHandler(cell -> cell.getRowIndex() >= 2, (cell, style) -> {
                    int rowIndex = cell.getRowIndex();
                    IndexedColors color = rowIndex % 2 == 0 ? IndexedColors.BLUE : IndexedColors.GREY_25_PERCENT;
                    style.setBackgroundColor(color);
                    return style;
                });
        // 数据写入excel
        writer.write(createTestListObject2());
    }

    private static int rowSize = 14;
    private static String[] PROP1_VALUES = {"aaa1", "aaa2", "aaa3"};
    private static MyConstraintEnum[] enValues = MyConstraintEnum.values();
    private static List<MyModel1> createTestListObject2() {
        List<MyModel1> list = new ArrayList<>();

        for (int i = 0; i < rowSize; i++) {
            MyModel1 ts = new MyModel1();
            ts.setPr1("pr1 == " + i);
            ts.setPr2(UUID.randomUUID().toString());
            int index = i + (Math.random() > 0.8 ? 1 : 0);
            ts.setPropValue2(enValues[index % enValues.length].getValue());
            ts.setPropValue1(PROP1_VALUES[index % PROP1_VALUES.length]);
            ts.setScore((int) (Math.random() * 100));
            ts.setDate(new Date());
            list.add(ts);
        }
        return list;
    }
}
