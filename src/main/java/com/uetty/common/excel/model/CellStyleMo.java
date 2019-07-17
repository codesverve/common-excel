package com.uetty.common.excel.model;

import org.apache.poi.ss.usermodel.*;

import java.util.Objects;

/**
 * @Author: Vince
 * @Date: 2019/7/16 14:21
 */
public class CellStyleMo implements Cloneable {

    HorizontalAlignment horizontalAlign;

    VerticalAlignment verticalAlign;

    BorderStyle borderStyle;

    IndexedColors borderColor;

    IndexedColors backgroundColor;

    Boolean wrapText;

    FontStyleMo fontStyle;

    Font font;

    public HorizontalAlignment getHorizontalAlign() {
        return horizontalAlign;
    }

    public void setHorizontalAlign(HorizontalAlignment horizontalAlign) {
        this.horizontalAlign = horizontalAlign;
    }

    public VerticalAlignment getVerticalAlign() {
        return verticalAlign;
    }

    public void setVerticalAlign(VerticalAlignment verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    public BorderStyle getBorderStyle() {
        return borderStyle;
    }

    public void setBorderStyle(BorderStyle borderStyle) {
        this.borderStyle = borderStyle;
    }

    public IndexedColors getBorderColor() {
        return borderColor;
    }

    public void setBorderColor(IndexedColors borderColor) {
        this.borderColor = borderColor;
    }

    public FontStyleMo getFontStyle() {
        return fontStyle;
    }

    public void setFontStyle(FontStyleMo fontStyle) {
        this.fontStyle = fontStyle;
    }

    public IndexedColors getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(IndexedColors backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    public Font getFont() {
        return font;
    }

    public void setFont(Font font) {
        this.font = font;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CellStyleMo that = (CellStyleMo) o;
        return horizontalAlign == that.horizontalAlign &&
                verticalAlign == that.verticalAlign &&
                borderStyle == that.borderStyle &&
                borderColor == that.borderColor &&
                backgroundColor == that.backgroundColor &&
                Objects.equals(wrapText, that.wrapText) &&
                Objects.equals(font, that.font);
    }

    @Override
    public int hashCode() {
        return Objects.hash(horizontalAlign, verticalAlign, borderStyle, borderColor, backgroundColor, wrapText, font);
    }

    @Override
    public CellStyleMo clone() {
        try {
            return (CellStyleMo) super.clone();
        } catch (CloneNotSupportedException e) {
            return null;
        }
    }
}
