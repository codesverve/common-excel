package com.uetty.common.excel.demo.headless;

import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.metadata.BaseRowModel;
import com.uetty.common.excel.anno.CellStyle;
import com.uetty.common.excel.anno.ColumnWidth;
import com.uetty.common.excel.anno.ExplicitConstraint;
import com.uetty.common.excel.anno.FontStyle;
import com.uetty.common.excel.constant.StyleType;
import com.uetty.common.excel.demo.MyConstraintEnum;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Date;

// 宽度默认值
@SuppressWarnings("unused")
@ColumnWidth(width = 50)
// 同时作用于表标题和内容的样式默认值
@CellStyle(backgroundColor = IndexedColors.CORAL, fontStyle = @FontStyle(color = IndexedColors.BLACK, size = 14))
public class MyModel2 extends BaseRowModel {

    // easyexcel 解析注解
    @ExcelColumnNum(value = 0)
    // 作用于该列内容的样式
    @CellStyle(type = StyleType.BODY_STYLE, backgroundColor = IndexedColors.GREY_80_PERCENT, fontStyle = @FontStyle(color = IndexedColors.LIGHT_BLUE))
    private String pr1;

    @ExcelColumnNum(value = 1)
    private String pr2;

    @ExcelColumnNum(value = 2)
    // 数组方式指定的下拉框约束
    @ExplicitConstraint(source = {"aaa1", "aaa2", "aaa3"})
    @ColumnWidth(width = 100)
    private String propValue1;

    @ExcelColumnNum(value = 3)
    @ExplicitConstraint(enumSource = MyConstraintEnum.class)
    private String propValue2;

    @ExcelColumnNum(value = 4)
    @ColumnWidth(width = 50)
    private Integer score;

    @ExcelColumnNum(value = 0, format = "yyyy-MM-dd HH:mm:ss")
    private Date date;

    public String getPr1() {
        return pr1;
    }

    void setPr1(String pr1) {
        this.pr1 = pr1;
    }

    public String getPr2() {
        return pr2;
    }

    void setPr2(String pr2) {
        this.pr2 = pr2;
    }

    public String getPropValue1() {
        return propValue1;
    }

    void setPropValue1(String propValue1) {
        this.propValue1 = propValue1;
    }

    public String getPropValue2() {
        return propValue2;
    }

    void setPropValue2(String propValue2) {
        this.propValue2 = propValue2;
    }

    public Integer getScore() {
        return score;
    }

    void setScore(Integer score) {
        this.score = score;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
