package com.uetty.common.excel.demo.withhead;

import com.uetty.common.excel.demo.MyConstraintEnum;
import com.uetty.common.excel.demo.MyModel1;
import com.uetty.common.excel.easyexcel.hssf.XlsExcelWriter;
import com.uetty.common.excel.util.ExcelHelper;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

/**
 * @Author: Vince
 * @Date: 2019/7/16 12:12
 */
public class TestWithHeadModel {

    public static void main(String[] args) throws IOException {
        XlsExcelWriter writer = new XlsExcelWriter("/data/test.xls", MyModel1.class, 0);

        // 合并单元格
        writer.addMergeRange(8, 9, 6, 6)
                // 另外添加下拉框约束
                .addExplicitConstraint(3, 4, 2, 4, new String[]{"理学", "文学", "工学"})
                // 自定义链式样式处理handler
                // 参数1：是否进入handler
                // 参数2：handler逻辑，入参cell和当前样式，出参样式类将作为下一个handler的入参
                .addCustomCellStyleHandler(cell -> cell.getRowIndex() < 2, (cell, style) -> {
                    style.setBackgroundColor(IndexedColors.CORAL);
                    return style;
                })
                .addCustomCellStyleHandler(cell -> cell.getRowIndex() >= 2, (cell, style) -> {
                    System.out.println(ExcelHelper.getCellValue(cell, null));
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
            ts.setPropValue1(PROP1_VALUES[index % PROP1_VALUES.length]);
            ts.setPropValue2(enValues[index % enValues.length].getValue());
            ts.setScore((int) (Math.random() * 100));
            ts.setDate(new Date());
            list.add(ts);
        }
        return list;
    }
}
