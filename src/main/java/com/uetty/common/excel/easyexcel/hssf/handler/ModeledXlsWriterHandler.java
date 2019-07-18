package com.uetty.common.excel.easyexcel.hssf.handler;

import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.uetty.common.excel.anno.CellStyle;
import com.uetty.common.excel.anno.*;
import com.uetty.common.excel.constant.ConstraintValue;
import com.uetty.common.excel.constant.NoneConstraint;
import com.uetty.common.excel.constant.StyleType;
import com.uetty.common.excel.model.*;
import com.uetty.common.excel.util.ReflectUtil;
import org.apache.poi.hssf.usermodel.HSSFDataValidationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.lang.reflect.Field;
import java.util.*;
import java.util.function.BiFunction;
import java.util.function.Predicate;
import java.util.stream.Collectors;

@SuppressWarnings("unused")
public class ModeledXlsWriterHandler implements WriteHandler {

    private static final int STANDARD_CHAR_WIDTH = 256; // 标准字符宽度

    private Class<? extends BaseRowModel> modelClazz;

    private SheetProperty sheetProperty;

    private Sheet sheet;

    private int startRow;
    // 是否有ExcelProperty注解
    private boolean hasExcelProperty = true;

    @SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
    private Map<CellStyleMo, org.apache.poi.ss.usermodel.CellStyle> cellStyleCached = new HashMap<>();
    @SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
    private Map<FontStyleMo, Font> fontStyleCached = new HashMap<>();

    private List<Field> fields = new ArrayList<>();

    public ModeledXlsWriterHandler(Class<? extends BaseRowModel> modelClazz, int startRow) {
        this.modelClazz = Objects.requireNonNull(modelClazz);
        this.startRow = startRow;
        initProperty();
    }

    private void initProperty() {
        this.sheetProperty = new SheetProperty();
        this.sheetProperty.setStartRow(this.startRow);
        CellFreeze cellFreeze = modelClazz.getAnnotation(CellFreeze.class);
        ColumnWidth columnWidth = modelClazz.getAnnotation(ColumnWidth.class);
        SheetStyleMo sheetStyleMo = resolveSheetStyle(cellFreeze, columnWidth);
        if (sheetStyleMo == null) {
            sheetStyleMo = new SheetStyleMo();
        }
        sheetProperty.setSheetStyle(sheetStyleMo);

        CellStyles cellStyles = modelClazz.getAnnotation(CellStyles.class);
        if (cellStyles != null) {
            CellStyle[] cellStyleArr = cellStyles.value();
            for (CellStyle cellStyle : cellStyleArr) {
                setGlobalCellStyle(sheetProperty, cellStyle);
            }
        } else {
            CellStyle cellStyle = modelClazz.getAnnotation(CellStyle.class);
            setGlobalCellStyle(sheetProperty, cellStyle);
        }

        // 字段列表
        List<Field> declaredFields = ReflectUtil.getDeclaredFields(modelClazz);
        fields = declaredFields.stream().filter(f -> {
            f.setAccessible(true);
            return f.getAnnotation(ExcelProperty.class) != null || f.getAnnotation(ExcelColumnNum.class) != null;
        }).sorted((f1, f2) -> {
            ExcelProperty prop1 = f1.getAnnotation(ExcelProperty.class);
            ExcelProperty prop2 = f2.getAnnotation(ExcelProperty.class);
            int index1 = prop1 != null ? prop1.index() : f1.getAnnotation(ExcelColumnNum.class).value();
            int index2 = prop2 != null ? prop2.index() : f2.getAnnotation(ExcelColumnNum.class).value();
            return index1 - index2;
        }).collect(Collectors.toList());

        hasExcelProperty = fields.stream().anyMatch(f -> f.getAnnotation(ExcelProperty.class) != null);

        // 标题行数
        int titleRowNum = fields.stream().map(f -> f.getAnnotation(ExcelProperty.class) != null ? f.getAnnotation(ExcelProperty.class).value().length : 0).max(Integer::compareTo).orElse(0);
        sheetProperty.setHeadRowNum(titleRowNum);

        for (int i = 0; i < fields.size(); i++) {
            readFieldAnnotation(i, fields.get(i));
        }
    }

    private SheetStyleMo resolveSheetStyle(CellFreeze cellFreeze, ColumnWidth columnWidth) {
        SheetStyleMo sheetStyleMo = new SheetStyleMo();
        if (cellFreeze != null) {
            sheetStyleMo.setFreezeCol(cellFreeze.freezeCol());
            sheetStyleMo.setFreezeRow(cellFreeze.freezeRow());
        }
        if (columnWidth != null) {
            sheetStyleMo.setDefaultColWidth(columnWidth.width());
        }
        return sheetStyleMo;
    }

    private void readFieldAnnotation(int index, Field field) {
        ColumnWidth columnStyle = field.getAnnotation(ColumnWidth.class);
        if (columnStyle != null) {
            sheetProperty.setColumnWidth(index, columnStyle.width());
        }

        CellStyles cellStyles = field.getAnnotation(CellStyles.class);
        if (cellStyles != null) {
            CellStyle[] cellStyleArr = cellStyles.value();
            for (CellStyle cellStyle : cellStyleArr) {
                setColumnCellStyle(sheetProperty, index, cellStyle);
            }
        } else {
            CellStyle cellStyle = field.getAnnotation(CellStyle.class);
            setColumnCellStyle(sheetProperty, index, cellStyle);
        }

        ExplicitConstraint explicitConstraint = field.getAnnotation(ExplicitConstraint.class);
        String[] explicitArray = resolveExplicitConstraint(explicitConstraint);
        if (explicitArray != null && explicitArray.length > 0) {
            sheetProperty.addExplicitConstraint(sheetProperty.getHeadRowNum() + startRow, Integer.MAX_VALUE, index, index, explicitArray);
        }
    }

    private void setColumnCellStyle(SheetProperty sheetProperty, int columnIndex, CellStyle cellStyle) {
        if (cellStyle == null) return;
        StyleType type = cellStyle.type();
        CellStyleMo cellStyleMo = resolveCellStyle(cellStyle);
        if (type == StyleType.HEAD_STYLE) {
            sheetProperty.setColumnHeaderStyle(columnIndex, cellStyleMo);
        } else if (type == StyleType.BODY_STYLE) {
            sheetProperty.setColumnBodyStyle(columnIndex, cellStyleMo);
        } else {
            sheetProperty.setColumnStyle(columnIndex, cellStyleMo);
        }
    }

    private String[] resolveExplicitConstraint(ExplicitConstraint explicitConstraint) {
        if (explicitConstraint == null) return null;
        String[] source = explicitConstraint.source();
        if (source.length > 0) {
            return source;
        }
        Class<? extends Enum<? extends ConstraintValue>> enumSource = explicitConstraint.enumSource();
        if (enumSource != NoneConstraint.class) {
            Enum<? extends ConstraintValue>[] constants = enumSource.getEnumConstants();
            source = new String[constants.length];
            for (int i = 0; i < constants.length; i++) {
                ConstraintValue cons = (ConstraintValue) constants[i];
                source[i] = cons.getValue();
            }
        }
        if (source.length > 0) {
            return source;
        }
        return null;
    }

    private void setGlobalCellStyle(SheetProperty sheetProperty, CellStyle cellStyle) {
        if (cellStyle == null) return;
        StyleType type = cellStyle.type();
        CellStyleMo cellStyleMo = resolveCellStyle(cellStyle);
        if (type == StyleType.HEAD_STYLE) {
            sheetProperty.setGlobalHeaderStyle(cellStyleMo);
        } else if (type == StyleType.BODY_STYLE) {
            sheetProperty.setGlobalBodyStyle(cellStyleMo);
        } else {
            sheetProperty.setGlobalCellStyle(cellStyleMo);
        }
    }

    private CellStyleMo resolveCellStyle(CellStyle cellStyle) {
        if (cellStyle == null) return null;
        CellStyleMo cellStyleMo = new CellStyleMo();
        cellStyleMo.setBackgroundColor(cellStyle.backgroundColor());
        cellStyleMo.setBorderColor(cellStyle.borderColor());
        cellStyleMo.setBorderStyle(cellStyle.borderStyle());
        cellStyleMo.setHorizontalAlign(cellStyle.horizontalAlign());
        cellStyleMo.setVerticalAlign(cellStyle.verticalAlign());
        cellStyleMo.setWrapText(cellStyle.wrapText());
        FontStyle fontStyle = cellStyle.fontStyle();
        FontStyleMo fontStyleMo = new FontStyleMo();
        fontStyleMo.setBold(fontStyle.bold());
        fontStyleMo.setColor(fontStyle.color());
        fontStyleMo.setName(fontStyle.name());
        fontStyleMo.setSize(fontStyle.size());
        cellStyleMo.setFontStyle(fontStyleMo);
        return cellStyleMo;
    }

    private void resolveAddress(CellRangeAddress address) {
        if (address.getFirstColumn() == Integer.MAX_VALUE) {
            address.setFirstColumn(-1);
        }
        if (address.getLastColumn() == Integer.MAX_VALUE) {
            address.setLastColumn(-1);
        }
        if (address.getFirstRow() == Integer.MAX_VALUE) {
            address.setFirstRow(-1);
        }
        if (address.getLastRow() == Integer.MAX_VALUE) {
            address.setLastRow(-1);
        }
    }

    @Override
    public void sheet(int sheetNo, Sheet sheet) {
        this.sheet = sheet;

        if (sheetProperty.getExplicitConstraints() != null && sheetProperty.getExplicitConstraints().size() > 0) {
            sheetProperty.getExplicitConstraints().forEach((address, explicitList) -> {
                if (explicitList == null || explicitList.length == 0) return;
                HSSFDataValidationHelper dvHelper = new HSSFDataValidationHelper((HSSFSheet) sheet);
                CellRangeAddressList rangeList = new CellRangeAddressList();
                CellRangeAddress addr = address.clone();
                resolveAddress(addr);
                rangeList.addCellRangeAddress(addr);
                DataValidationConstraint constraint = dvHelper.createExplicitListConstraint(explicitList);
                DataValidation validation = dvHelper.createValidation(constraint, rangeList);
                sheet.addValidationData(validation);
            });
        }

        if (sheetProperty.getSheetStyle() != null) {
            SheetStyleMo sheetStyle = sheetProperty.getSheetStyle();
            if (sheetStyle.getFreezeCol() > 0 || sheetStyle.getFreezeRow() > 0) {
                sheet.createFreezePane(sheetStyle.getFreezeCol(), sheetStyle.getFreezeRow());
            }
        }

        for (int col = 0; col < fields.size(); col++) {
            double drawColumnWidth = sheetProperty.getDrawColumnWidth(col);
            if (drawColumnWidth >= 0) {
                sheet.setColumnWidth(col, (int) (drawColumnWidth * STANDARD_CHAR_WIDTH));
            }
        }

        List<ICellRange> mergeRanges = sheetProperty.getMergeRanges();
        for (ICellRange mergeRange : mergeRanges) {
            sheet.addMergedRegion(mergeRange);
        }
    }

    @Override
    public void row(int rowIndex, Row row) {
    }


    @Override
    public void cell(int colIndex, Cell cell) {
        CellStyleMo style = sheetProperty.getDrawCellStyle(cell);

        org.apache.poi.ss.usermodel.CellStyle cellStyle = getCellStyle(style);
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    private org.apache.poi.ss.usermodel.CellStyle getCellStyle(CellStyleMo styleMo) {
        if (styleMo == null) return null;

        // 设置字体
        Font font = fontStyleCached.get(styleMo.getFontStyle());
        if (font == null) {
            font = sheet.getWorkbook().createFont();
            font.setFontName(styleMo.getFontStyle().getName());
            font.setColor(styleMo.getFontStyle().getColor().getIndex());
            if (styleMo.getFontStyle().getSize() != -1) {
                font.setFontHeightInPoints((short) styleMo.getFontStyle().getSize());
            }
        }
        styleMo.setFont(font);

        org.apache.poi.ss.usermodel.CellStyle cellStyle = cellStyleCached.get(styleMo);
        if (cellStyle == null) {
            cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.setFont(styleMo.getFont());
            cellStyle.setVerticalAlignment(styleMo.getVerticalAlign());
            cellStyle.setAlignment(styleMo.getHorizontalAlign());
            cellStyle.setLocked(true);
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // 背景颜色
            cellStyle.setFillBackgroundColor(styleMo.getBackgroundColor().index);
            cellStyle.setFillForegroundColor(styleMo.getBackgroundColor().index);

            cellStyle.setBorderBottom(styleMo.getBorderStyle());
            cellStyle.setBorderLeft(styleMo.getBorderStyle());
            cellStyle.setBorderRight(styleMo.getBorderStyle());
            cellStyle.setBorderTop(styleMo.getBorderStyle());

            cellStyle.setBottomBorderColor(styleMo.getBorderColor().getIndex());
            cellStyle.setLeftBorderColor(styleMo.getBorderColor().getIndex());
            cellStyle.setRightBorderColor(styleMo.getBorderColor().getIndex());
            cellStyle.setTopBorderColor(styleMo.getBorderColor().getIndex());

            cellStyle.setWrapText(styleMo.getWrapText());
        }

        return cellStyle;
    }

    public boolean hasExcelProperty() {
        return hasExcelProperty;
    }

    public int getStartRow() {
        return startRow;
    }
    public void addMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.sheetProperty.addMergeRange(firstRow, lastRow, firstCol, lastCol);
    }
    public boolean addMergeRangeNoneException(int firstRow, int lastRow, int firstCol, int lastCol) {
        return this.sheetProperty.addMergeRangeNoneException(firstRow, lastRow, firstCol, lastCol);
    }
    public void removeMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.sheetProperty.removeMergeRange(firstRow, lastRow, firstCol, lastCol);
    }
    public void clearMergeRange() {
        this.sheetProperty.clearMergeRange();
    }

    public void setGlobalCellStyle(CellStyleMo globalCellStyle) {
        this.sheetProperty.setGlobalCellStyle(globalCellStyle);
    }
    public void setGlobalBodyStyle(CellStyleMo globalBodyStyle) {
        this.sheetProperty.setGlobalBodyStyle(globalBodyStyle);
    }

    public int getHeadRowNum() {
        return this.sheetProperty.getHeadRowNum();
    }

    public void setSheetStyle(SheetStyleMo sheetStyle) {
        this.sheetProperty.setSheetStyle(sheetStyle);
    }
    public void setColumnStyle(int column, CellStyleMo style) {
        this.sheetProperty.setColumnStyle(column, style);
    }
    public CellStyleMo getColumnStyle(int column) {
        return this.sheetProperty.getColumnStyle(column);
    }

    public void setColumnHeaderStyle(int column, CellStyleMo style) {
        this.sheetProperty.setColumnHeaderStyle(column, style);
    }
    public CellStyleMo getColumnHeaderStyle(int column) {
        return this.sheetProperty.getColumnHeaderStyle(column);
    }
    public void clearColumnHeaderStyles() {
        this.sheetProperty.clearColumnHeaderStyles();
    }
    public void removeColumnHeaderStyle(int column) {
        this.sheetProperty.removeColumnHeaderStyle(column);
    }

    public void clearColumnBodyStyles() {
        this.sheetProperty.clearColumnBodyStyles();
    }
    public CellStyleMo getColumnBodyStyle(int column) {
        return this.sheetProperty.getColumnBodyStyle(column);
    }
    public void setColumnBodyStyle(int column, CellStyleMo style) {
        this.sheetProperty.setColumnBodyStyle(column, style);
    }
    public void removeColumnBodyStyle(int column) {
        this.sheetProperty.removeColumnBodyStyle(column);
    }

    public void clearColumnWidths() {
        this.sheetProperty.clearColumnWidths();
    }
    public Double getColumnWidth(int column) {
        return this.sheetProperty.getColumnWidth(column);
    }
    public void setColumnWidth(int column, int width) {
        this.sheetProperty.setColumnWidth(column, width);
    }
    public void removeColumnWidth(int column) {
        this.sheetProperty.removeColumnWidth(column);
    }

    /**
     * 给指定区域添加下拉列表约束
     */
    public void addExplicitConstraint(int firstRow, int lastRow, int firstCol, int lastCol, String[] explicitArray) {
        this.sheetProperty.addExplicitConstraint(firstRow, lastRow, firstCol, lastCol, explicitArray);
    }
    public void removeRangeConstraint(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.sheetProperty.removeRangeConstraint(firstRow, lastRow, firstCol, lastCol);
    }
    public void clearExplicitConstraints() {
        this.sheetProperty.clearExplicitConstraints();
    }
    public String[] getExplicitConstraint(int firstRow, int lastRow, int firstCol, int lastCol) {
        return this.sheetProperty.getExplicitConstraint(firstRow, lastRow, firstCol, lastCol);
    }

    public void clearCustomCellStyleHandlers() {
        this.sheetProperty.clearCustomCellStyleHandlers();
    }
    public void removeCustomCellStyleHandler(Predicate<Cell> predicate) {
        this.sheetProperty.removeCustomCellStyleHandler(predicate);
    }
    public void addCustomCellStyleHandler(Predicate<Cell> predicate, BiFunction<Cell, CellStyleMo, CellStyleMo> handler) {
        this.sheetProperty.addCustomCellStyleHandler(predicate, handler);
    }

    public void setFreeze(int freezeCol, int freezeRow) {
        this.sheetProperty.setFreeze(freezeCol,freezeRow);
    }
}
