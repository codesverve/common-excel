package com.uetty.common.excel.model;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Objects;

/**
 * @Author: Vince
 * @Date: 2019/7/16 14:23
 */
public class FontStyleMo {

    String name;

    double size;

    IndexedColors color;

    boolean bold;

    Font font;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public double getSize() {
        return size;
    }

    public void setSize(double size) {
        this.size = size;
    }

    public IndexedColors getColor() {
        return color;
    }

    public void setColor(IndexedColors color) {
        this.color = color;
    }

    public boolean getBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
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
        FontStyleMo that = (FontStyleMo) o;
        return Double.compare(that.size, size) == 0 &&
                bold == that.bold &&
                Objects.equals(name, that.name) &&
                color == that.color;
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, size, color, bold);
    }
}
