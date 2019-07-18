package com.uetty.common.excel.model;

import org.apache.poi.hssf.record.RecordInputStream;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@SuppressWarnings({"WeakerAccess", "unused"})
public class ICellRange extends CellRangeAddress implements Cloneable {

    public ICellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        super(firstRow, lastRow, firstCol, lastCol);
    }

    public ICellRange(RecordInputStream in) {
        super(in);
    }

    /**
     * 挖除区域
     */
    public List<ICellRange> subtractRange(ICellRange range) {
        List<ICellRange> list = new ArrayList<>();
        if (!isOverlap(range)) {
            list.add(clone());
        } else {
            if (getMinColumn() < range.getMinColumn()) {
                int top = getMinRow();
                int bottom = getMaxRow();
                int left = getMinColumn();
                int right = Math.min(getMaxColumn(), range.getMinColumn() - 1);
                ICellRange urange = new ICellRange(top, bottom, left, right);
                list.add(urange);
            }
            if (getMaxColumn() > range.getMaxColumn()) {
                int top = getMinRow();
                int bottom = getMaxRow();
                int left = Math.max(getMinColumn(), range.getMaxColumn() + 1);
                int right = getMaxColumn();
                ICellRange urange = new ICellRange(top, bottom, left, right);
                list.add(urange);
            }

            int left = Math.max(getMinColumn(), range.getMinColumn());
            int right = Math.min(getMaxColumn(), range.getMaxColumn());
            if (right >= left) {
                if (getMinRow() < range.getMinRow()) {
                    int top = getMinRow();
                    int bottom = Math.min(getMaxRow(), range.getMinRow() - 1);
                    ICellRange urange = new ICellRange(top, bottom, left, right);
                    list.add(urange);
                }

                if (getMaxRow() > range.getMaxRow()) {
                    int top = Math.max(getMinRow(), range.getMaxRow() + 1);
                    int bottom = getMaxRow();
                    ICellRange urange = new ICellRange(top, bottom, left, right);
                    list.add(urange);
                }
            }
        }
        return list;
    }

    /**
     * 返回重叠区域
     */
    public ICellRange overlapRange(ICellRange range) {
        int left = Math.max(getMinColumn(), range.getMinColumn());
        int right = Math.min(getMaxColumn(), range.getMaxColumn());
        if (left > right) return null;
        int top = Math.max(getMinRow(), range.getMinRow());
        int bottom = Math.min(getMaxRow(), range.getMaxRow());
        if (top > bottom) return null;
        return new ICellRange(top, bottom, left, right);
    }

    /**
     * 是否重叠
     */
    public boolean isOverlap(ICellRange range) {
        int top = Math.max(getMinRow(), range.getMinRow());
        int bottom = Math.min(getMaxRow(), range.getMaxRow());
        if (top > bottom) return false;
        int left = Math.max(getMinColumn(), range.getMinColumn());
        int right = Math.min(getMaxColumn(), range.getMaxColumn());
        return left <= right;
    }

    public String toString1() {
        return getMinRow() + " - " + getMaxRow() + " - " + getMinColumn() + " - " + getMaxColumn();
    }


    @Override
    public ICellRange clone() {
        try {
            return (ICellRange) super.clone();
        } catch (CloneNotSupportedException e) {
            return null;
        }
    }

    public static void main(String[] args) {
        ICellRange r1 = new ICellRange(2, 5, 1, 4);
        ICellRange r2 = new ICellRange(6, 8, 2, 5);
        ICellRange r3 = new ICellRange(1, 3, 3, 6);
        ICellRange r4 = new ICellRange(3, 4, 2, 3);

        System.out.println(r1.isOverlap(r2));
        System.out.println(r1.isOverlap(r3));
        System.out.println(r1.isOverlap(r4));
        System.out.println(r4.isOverlap(r1));
        System.out.println();

        System.out.println(r1.overlapRange(r2));
        System.out.println(r1.overlapRange(r3).toString1());
        System.out.println(r1.overlapRange(r4).toString1());
        System.out.println(r4.overlapRange(r1).toString1());
        System.out.println();

        System.out.println(r1.subtractRange(r2).stream().map(ICellRange::toString1).collect(Collectors.toList()));
        System.out.println(r1.subtractRange(r3).stream().map(ICellRange::toString1).collect(Collectors.toList()));
        System.out.println(r1.subtractRange(r4).stream().map(ICellRange::toString1).collect(Collectors.toList()));
        System.out.println(r4.subtractRange(r1).stream().map(ICellRange::toString1).collect(Collectors.toList()));
    }
}
