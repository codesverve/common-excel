package com.uetty.common.excel.demo.multisheet;

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

public class TestMultipleSheet {
    public static void main(String[] args) throws IOException {
        XlsExcelWriter writer = new XlsExcelWriter("/data/com-exc4.xls");

        writer.addNewSheet(MyModel1.class, 0);
        // 合并单元格
        writer.addMergeRange(8, 9, 6, 6)
                // 另外添加下拉框约束
//                .addExplicitConstraint(3, 4, 2, 4, new String[]{"理学", "文学", "工学"})
                // 自定义链式样式处理handler
                // 参数1：是否进入handler
                // 参数2：handler逻辑，入参cell和当前样式，出参样式类将作为下一个handler的入参
                .addCustomCellStyleHandler(cell -> cell.getRowIndex() < 2, (cell, style) -> {
                    style.setBackgroundColor(IndexedColors.SEA_GREEN);
                    return style;
                })
                .addCustomCellStyleHandler(cell -> cell.getRowIndex() >= 2, (cell, style) -> {
                    System.out.println(ExcelHelper.getCellValue(cell, null));
                    int rowIndex = cell.getRowIndex();
                    IndexedColors color = rowIndex % 2 == 0 ? IndexedColors.WHITE : IndexedColors.GREY_40_PERCENT;
                    style.setBackgroundColor(color);
                    return style;
                });
        // 数据写入excel
        writer.write(createTestListObject2());

        writer.addNewSheet();
        writer.write(createTestListObject2());

        writer.addNewSheet(MyModel1.class, 0);
        writer.setSheetName("sheet third");
//        writer.addCustomCellStyleHandler()
        writer.write(createTestListObject2());

        writer.flush();
    }

    private static final int ROW_SIZE = 14;
    private static String[] PROP1_VALUES = {"aaa1", "aaa2", "aaa3"};
    private static MyConstraintEnum[] enValues = MyConstraintEnum.values();
    private static List<MyModel1> createTestListObject2() {
        List<MyModel1> list = new ArrayList<>();

        for (int i = 0; i < ROW_SIZE; i++) {
            MyModel1 ts = new MyModel1();
            int index = i + (Math.random() > 0.8 ? 1 : 0);
            ts.setPr1("pr1 == " + i);
            ts.setPr2(UUID.randomUUID().toString());
            ts.setPropValue1(PROP1_VALUES[index % PROP1_VALUES.length]);
            ts.setScore((int) (Math.random() * 100));
            ts.setPropValue2(enValues[index % enValues.length].getValue());
            ts.setDate(new Date());
            list.add(ts);
        }
        return list;
    }
}
