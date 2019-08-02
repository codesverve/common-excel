package com.uetty.common.excel.anno;

import com.uetty.common.excel.constant.StyleType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(CellStyles.class)
public @interface CellStyle {

    StyleType type() default StyleType.ALL_STYLE; // 类型: 标题、内容、全部

    /**
     * 水平居中方式 默认左居中
     * @see HorizontalAlignment
     * @return : org.apache.poi.ss.usermodel.HorizontalAlignment
     * @author : Vince
     */
    HorizontalAlignment horizontalAlign() default HorizontalAlignment.LEFT;

    /**
     * 垂直居中方式 默认居中
     * @see VerticalAlignment
     * @return : org.apache.poi.ss.usermodel.VerticalAlignment
     * @author : Vince
     */
    VerticalAlignment verticalAlign() default VerticalAlignment.CENTER;


    /**
     * 边框方式 默认无
     * @return : org.apache.poi.ss.usermodel.BorderStyle
     * @author : Vince
     */
    BorderStyle borderStyle() default BorderStyle.NONE;


    /**
     * 边框颜色 默认白
     * @return : org.apache.poi.ss.usermodel.IndexedColors
     * @author : Vince
     */
    IndexedColors borderColor() default IndexedColors.AUTOMATIC;


    /**
     * 字体设置
     * @see org.apache.poi.xssf.usermodel.XSSFFont
     * @see org.apache.poi.hssf.usermodel.HSSFFont
     * @return : com.uetty.common.excel.anno.FontStyle
     * @author : Vince
     */
    FontStyle fontStyle() default @FontStyle();


    /**
     * 背景颜色
     * @return : org.apache.poi.ss.usermodel.IndexedColors
     * @author : Vince
     */
    IndexedColors backgroundColor() default IndexedColors.WHITE;


    /**
     * 单元格是否换行
     * @return : boolean
     * @author : Vince
     */
    boolean wrapText() default false;

}
