package com.uetty.common.excel.demo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import com.uetty.common.excel.anno.CellStyle;
import com.uetty.common.excel.anno.ColumnWidth;
import com.uetty.common.excel.anno.ExplicitConstraint;
import com.uetty.common.excel.anno.FontStyle;
import com.uetty.common.excel.constant.StyleType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Date;

/**
 * @Author: Vince
 * @Date: 2019/7/16 12:09
 */
// 宽度默认值
@ColumnWidth(width = 40)
// 同时作用于表标题和内容的样式默认值
@CellStyle(borderColor = IndexedColors.DARK_TEAL, backgroundColor = IndexedColors.CORAL, fontStyle = @FontStyle(color = IndexedColors.ORANGE, size = 14))
public class MyModel1 extends BaseRowModel {

    // easyexcel 解析注解
    @ExcelProperty(value={"pri", "pr1"}, index = 0)
    // 作用于该列标题的样式
    @CellStyle(type = StyleType.HEAD_STYLE, borderColor = IndexedColors.GOLD, borderStyle = BorderStyle.DASHED)
    // 作用于该列内容的样式
    @CellStyle(type = StyleType.BODY_STYLE, backgroundColor = IndexedColors.PINK, fontStyle = @FontStyle(color = IndexedColors.LIGHT_BLUE))
    String pr1;

    @ExcelProperty(value={"pri", "pr2"}, index = 1)
    // 作用于该列标题和内容的样式
    @CellStyle(borderColor = IndexedColors.GOLD, borderStyle = BorderStyle.DASHED)
    String pr2;

    @ExcelProperty(value = {"propValue1"}, index = 2)
    // 数组方式指定的下拉框约束
    @ExplicitConstraint(source = {"aaa1", "aaa2", "aaa3"})
    @ColumnWidth(width = 100)
    String propValue1;

    @ExcelProperty(value = {"propValue2"}, index = 3)
    @ExplicitConstraint(enumSource = MyConstraintEnum.class)
    String propValue2;

    @ExcelProperty(value = {"score"}, index = 4)
    @ColumnWidth(width = 50)
    Integer score;

    @ExcelProperty(value = {"date"}, index = 5, format = "yyyy-MM-dd HH:mm:ss")
    Date date;

    public String getPr1() {
        return pr1;
    }

    public void setPr1(String pr1) {
        this.pr1 = pr1;
    }

    public String getPr2() {
        return pr2;
    }

    public void setPr2(String pr2) {
        this.pr2 = pr2;
    }

    public String getPropValue1() {
        return propValue1;
    }

    public void setPropValue1(String propValue1) {
        this.propValue1 = propValue1;
    }

    public String getPropValue2() {
        return propValue2;
    }

    public void setPropValue2(String propValue2) {
        this.propValue2 = propValue2;
    }

    public Integer getScore() {
        return score;
    }

    public void setScore(Integer score) {
        this.score = score;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
