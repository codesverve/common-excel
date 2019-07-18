package com.uetty.common.excel.model;

import org.apache.poi.ss.usermodel.Cell;

import java.util.*;
import java.util.function.BiFunction;
import java.util.function.Predicate;

@SuppressWarnings("unused")
public class SheetProperty {

    private int startRow;
    private int headRowNum;
    private boolean needHead = true;
    private CellStyleMo globalCellStyle; // 单元格样式
    private CellStyleMo globalHeaderStyle; // 表头单元格样式
    private CellStyleMo globalBodyStyle; // 表数据单元格样式
    private SheetStyleMo sheetStyle;
    private List<ICellRange> mergeRanges = new ArrayList<>(); // 合并单元格
    private Map<Integer, CellStyleMo> columnStyles = new HashMap<>(); // 指定列单元格样式
    private Map<Integer, CellStyleMo> columnHeaderStyles = new HashMap<>(); // 指定列表头单元格样式
    private Map<Integer, CellStyleMo> columnBodyStyles = new HashMap<>(); // 指定列表数据单元格样式
    private Map<Integer, Double> columnWidths = new HashMap<>(); // 列宽度
    private Map<ICellRange, String[]> explicitConstraints = new HashMap<>(); // 简单下拉菜单约束
    // 单元格样式处理器链表，链式串联调用
    // Predicate用于判断是否进入该样式处理器，
    // BiFunction<Cell, CellStyleMo, CellStyleMo>处理样式，前两个参数为入参，最后的参数为出参，
    // 链表上一级的出参作为下一级的入参传入，最终出参作用于excel表单上
    private LinkedHashMap<Predicate<Cell>, BiFunction<Cell, CellStyleMo, CellStyleMo>> cellStyleHandlersLink = new LinkedHashMap<>();

    public int getStartRow() {
        return startRow;
    }
    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }
    public int getHeadRowNum() {
        return headRowNum;
    }
    public void setHeadRowNum(int headRowNum) {
        this.headRowNum = headRowNum;
    }
    public boolean isNeedHead() {
        return needHead;
    }
    public void setNeedHead(boolean needHead) {
        this.needHead = needHead;
    }
    public CellStyleMo getGlobalCellStyle() {
        return globalCellStyle;
    }
    public void setGlobalCellStyle(CellStyleMo globalCellStyle) {
        this.globalCellStyle = globalCellStyle;
    }
    public CellStyleMo getGlobalHeaderStyle() {
        return globalHeaderStyle;
    }

    public void setGlobalHeaderStyle(CellStyleMo globalHeaderStyle) {
        this.globalHeaderStyle = globalHeaderStyle;
    }

    public CellStyleMo getGlobalBodyStyle() {
        return globalBodyStyle;
    }

    public void setGlobalBodyStyle(CellStyleMo globalBodyStyle) {
        this.globalBodyStyle = globalBodyStyle;
    }

    public SheetStyleMo getSheetStyle() {
        return sheetStyle;
    }

    public void setSheetStyle(SheetStyleMo sheetStyle) {
        this.sheetStyle = sheetStyle;
    }

    public void addMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        ICellRange range = new ICellRange(firstRow, lastRow, firstCol, lastCol);
        for (ICellRange area : mergeRanges) {
            if (area.isOverlap(range)) {
                throw new RuntimeException("area [" + area + "] has bean merged");
            }
        }
        mergeRanges.add(range);
    }
    public boolean addMergeRangeNoneException(int firstRow, int lastRow, int firstCol, int lastCol) {
        ICellRange range = new ICellRange(firstRow, lastRow, firstCol, lastCol);
        for (ICellRange area : mergeRanges) {
            if (area.isOverlap(range)) {
                return false;
            }
        }
        return mergeRanges.add(range);
    }
    public void removeMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        ICellRange range = new ICellRange(firstRow, lastRow, firstCol, lastCol);
        mergeRanges.remove(range);
    }
    public void clearMergeRange() {
        mergeRanges.clear();
    }
    public List<ICellRange> getMergeRanges() {
        return mergeRanges;
    }

    public void setColumnStyle(int column, CellStyleMo style) {
        columnStyles.put(column, style);
    }
    public CellStyleMo getColumnStyle(int column) {
        return columnStyles.get(column);
    }
    public void clearColumnStyles() {
        columnStyles.clear();
    }
    public Map<Integer, CellStyleMo> getColumnStyles() {
        return columnStyles;
    }
    public void removeColumnStyle(int column) {
        columnStyles.remove(column);
    }

    public Map<Integer, CellStyleMo> getColumnHeaderStyles() {
        return columnHeaderStyles;
    }
    public void setColumnHeaderStyle(int column, CellStyleMo style) {
        columnHeaderStyles.put(column, style);
    }
    public CellStyleMo getColumnHeaderStyle(int column) {
        return columnHeaderStyles.get(column);
    }
    public void clearColumnHeaderStyles() {
        columnHeaderStyles.clear();
    }
    public void removeColumnHeaderStyle(int column) {
        columnHeaderStyles.remove(column);
    }

    public Map<Integer, CellStyleMo> getColumnBodyStyles() {
        return columnBodyStyles;
    }
    public void clearColumnBodyStyles() {
        columnBodyStyles.clear();
    }
    public CellStyleMo getColumnBodyStyle(int column) {
        return columnBodyStyles.get(column);
    }
    public void setColumnBodyStyle(int column, CellStyleMo style) {
        columnBodyStyles.put(column, style);
    }
    public void removeColumnBodyStyle(int column) {
        columnBodyStyles.remove(column);
    }

    public Map<Integer, Double> getColumnWidths() {
        return columnWidths;
    }
    public void clearColumnWidths() {
        columnWidths.clear();
    }
    public Double getColumnWidth(int column) {
        return columnWidths.get(column);
    }
    public void setColumnWidth(int column, double width) {
        columnWidths.put(column, width);
    }
    public void removeColumnWidth(int column) {
        columnWidths.remove(column);
    }

    public Map<ICellRange, String[]> getExplicitConstraints() {
        return new HashMap<>(explicitConstraints);
    }
    /**
     * 给指定区域添加下拉列表约束
     */
    public void addExplicitConstraint(int firstRow, int lastRow, int firstCol, int lastCol, String[] explicitArray) {
        if (explicitArray == null || explicitArray.length == 0) return;
        removeMergeRange(firstRow, lastRow, firstCol, lastCol);
        ICellRange range = new ICellRange(firstRow, lastRow, firstCol, lastCol);
        // 增加最新区域约束
        explicitConstraints.put(range, explicitArray);
    }
    public void removeRangeConstraint(int firstRow, int lastRow, int firstCol, int lastCol) {
        ICellRange range = new ICellRange(firstRow, lastRow, firstCol, lastCol);
        Map<ICellRange, String[]> overlapRanges = new HashMap<>();
        // 找出存在重叠区域的约束
        explicitConstraints.forEach((iCellRange, cons) -> {
            if (iCellRange.isOverlap(range)) {
                overlapRanges.put(iCellRange, cons);
            }
        });
        // 移除存在重叠区域的约束
        for (ICellRange iCellRange : overlapRanges.keySet()) {
            explicitConstraints.remove(iCellRange);
        }
        // 减掉重叠区域，加回剩余区域
        overlapRanges.forEach((iCellRange, cons) -> {
            List<ICellRange> iCellRanges = iCellRange.subtractRange(range);
            for (ICellRange cellRange : iCellRanges) {
                explicitConstraints.put(cellRange, cons);
            }
        });
    }
    public void clearExplicitConstraints() {
        explicitConstraints.clear();
    }
    public String[] getExplicitConstraint(int firstRow, int lastRow, int firstCol, int lastCol) {
        ICellRange range = new ICellRange(firstRow, lastRow, firstCol, lastCol);
        return explicitConstraints.get(range);
    }

    public LinkedHashMap<Predicate<Cell>, BiFunction<Cell, CellStyleMo, CellStyleMo>> getCellStyleHandlersLink() {
        return cellStyleHandlersLink;
    }
    public void clearCustomCellStyleHandlers() {
        cellStyleHandlersLink.clear();
    }
    public void removeCustomCellStyleHandler(Predicate<Cell> predicate) {
        cellStyleHandlersLink.remove(predicate);
    }
    public void addCustomCellStyleHandler(Predicate<Cell> predicate, BiFunction<Cell, CellStyleMo, CellStyleMo> handler) {
        cellStyleHandlersLink.put(predicate, handler);
    }

    // --------------------------------------------------- 上面excel工作前设置
    // --------------------------------------------------- 下面excel工作时使用

    public CellStyleMo getDrawCellStyle(Cell cell) {
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();

        boolean isHeader = needHead && rowIndex - startRow < headRowNum;

        CellStyleMo style;
        if (isHeader) { // withhead
            style = columnHeaderStyles.get(columnIndex);
        } else { // body
            style = columnBodyStyles.get(columnIndex);
        }
        if (style == null) {
            style = columnStyles.get(columnIndex);
        }
        if (style == null) {
            if (isHeader) { // withhead
                style = globalHeaderStyle;
            } else { // body
                style = globalBodyStyle;
            }
        }
        if (style == null) {
            style = globalCellStyle;
        }
        if (style != null) {
            style = style.clone(); // 防止污染到公共数据
        }
        Set<Predicate<Cell>> predicates = cellStyleHandlersLink.keySet();
        for (Predicate<Cell> predicate : predicates) {
            try {
                boolean test = predicate.test(cell);
                if (!test) continue;

                BiFunction<Cell, CellStyleMo, CellStyleMo> executor = cellStyleHandlersLink.get(predicate);
                style = executor.apply(cell,style);
            } catch (Exception e){
                e.printStackTrace();
            }
        }
        return style;
    }

    public double getDrawColumnWidth(int columnIndex) {
        Double width = columnWidths.get(columnIndex);
        if (width != null) {
            return width;
        }
        return sheetStyle != null && sheetStyle.getDefaultColWidth() != -1 ? sheetStyle.getDefaultColWidth() : -1;
    }

    public void setFreeze(int freezeCol, int freezeRow) {
        if (this.sheetStyle == null) {
            this.sheetStyle = new SheetStyleMo();
        }
        this.sheetStyle.setFreezeRow(freezeRow);
        this.sheetStyle.setFreezeCol(freezeCol);
    }
}
