package com.uetty.common.excel.model;

public class SheetStyleMo {

    private Integer freezeCol = 0;

    private Integer freezeRow = 0;

//    double defaultRowHeight = -1; // 高度被easyexcel限死掉了，没法设置

    private double defaultColWidth = -1 ;

    public Integer getFreezeCol() {
        return freezeCol;
    }

    public void setFreezeCol(Integer freezeCol) {
        this.freezeCol = freezeCol;
    }

    public Integer getFreezeRow() {
        return freezeRow;
    }

    public void setFreezeRow(Integer freezeRow) {
        this.freezeRow = freezeRow;
    }

//    public double getDefaultRowHeight() {
//        return defaultRowHeight;
//    }
//
//    public void setDefaultRowHeight(double defaultRowHeight) {
//        this.defaultRowHeight = defaultRowHeight;
//    }

    double getDefaultColWidth() {
        return defaultColWidth;
    }

    public void setDefaultColWidth(double defaultColWidth) {
        this.defaultColWidth = defaultColWidth;
    }
}
