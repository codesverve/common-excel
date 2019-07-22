package com.uetty.common.excel.easyexcel.hssf;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.support.ExcelTypeEnum;
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

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.function.BiFunction;
import java.util.function.Predicate;
import java.util.stream.Collectors;

@SuppressWarnings({"unused", "UnusedReturnValue", "WeakerAccess"})
public class XlsExcelWriter {

    private static final int STANDARD_CHAR_WIDTH = 256; // 标准字符宽度

    private Class<? extends BaseRowModel> modelClazz;

    private SheetProperty sheetProperty;

    private int startRow;
    // 是否有ExcelProperty注解
    private boolean hasExcelProperty = true;

    private List<Field> fields = new ArrayList<>();

    private String outputPath;
    private boolean needHead = true;
    private int sheetIndex = 1;
    private String sheetName;

    public XlsExcelWriter(String outputPath, Class<? extends BaseRowModel> modelClazz, int startRow) {
        this.outputPath = Objects.requireNonNull(outputPath);
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

    public void write(List<? extends BaseRowModel> list) throws IOException {
//        if (!hasExcelProperty) {
//            setNeedHead(false);
//        }

        File file = new File(outputPath);
        try (FileOutputStream fos = new FileOutputStream(file)) {

            ExcelWriter writer = EasyExcelFactory.getWriterWithTempAndHandler(null, fos, ExcelTypeEnum.XLS, getNeedHead(), new EasyExcelHandler());

            com.alibaba.excel.metadata.Sheet sheet = new com.alibaba.excel.metadata.Sheet(sheetIndex, 0, modelClazz);
            int iStartRow = this.startRow;
            if (sheetProperty.getHeadRowNum() == 0 || !getNeedHead()) { // easyexcel 本身在这一块有个小bug，兼容一下
                iStartRow = iStartRow - 1;
            }
            sheet.setStartRow(iStartRow);
            if (sheetName != null) {
                sheet.setSheetName(sheetName);
            } else {
                sheet.setSheetName("sheet-" + sheetIndex);
            }

            writer.write(list, sheet);
            //关闭资源
            writer.finish();
        }
    }

    public class EasyExcelHandler implements WriteHandler {
        private Sheet sheet;
        @SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
        private Map<CellStyleMo, org.apache.poi.ss.usermodel.CellStyle> cellStyleCached = new HashMap<>();
        @SuppressWarnings("MismatchedQueryAndUpdateOfCollection")
        private Map<FontStyleMo, Font> fontStyleCached = new HashMap<>();

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
                FontStyleMo fontStyle = styleMo.getFontStyle();
                font.setFontName(fontStyle.getName());
                font.setColor(fontStyle.getColor().getIndex());
                font.setBold(fontStyle.getBold());
                if (fontStyle.getSize() != -1) {
                    font.setFontHeightInPoints((short) fontStyle.getSize());
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
    }

    public boolean getNeedHead() {
        return needHead && hasExcelProperty;
    }
    public XlsExcelWriter setNeedHead(boolean needHead) {
        this.needHead = needHead && hasExcelProperty; // 没有含标题名称的注解，就不需要标题了
        return this;
    }
    public int getSheetIndex() {
        return sheetIndex;
    }
    public XlsExcelWriter setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
        return this;
    }
    public String getSheetName() {
        return sheetName;
    }
    public XlsExcelWriter setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }
    public XlsExcelWriter addMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.sheetProperty.addMergeRange(firstRow, lastRow, firstCol, lastCol);
        return this;
    }
    public XlsExcelWriter removeMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.sheetProperty.removeMergeRange(firstRow, lastRow, firstCol, lastCol);
        return this;
    }
    public XlsExcelWriter clearMergeRange() {
        this.sheetProperty.clearMergeRange();
        return this;
    }
    public XlsExcelWriter addExplicitConstraint(int firstRow, int lastRow, int firstCol, int lastCol, String[] explicitArray) {
        this.sheetProperty.addExplicitConstraint(firstRow, lastRow, firstCol, lastCol, explicitArray);
        return this;
    }
    public XlsExcelWriter removeRangeConstraint(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.sheetProperty.removeRangeConstraint(firstRow, lastRow, firstCol, lastCol);
        return this;
    }
    public XlsExcelWriter clearExplicitConstraints() {
        this.sheetProperty.clearExplicitConstraints();
        return this;
    }
    public XlsExcelWriter clearCustomCellStyleHandlers() {
        this.sheetProperty.clearCustomCellStyleHandlers();
        return this;
    }
    public XlsExcelWriter removeCustomCellStyleHandler(Predicate<Cell> predicate) {
        this.sheetProperty.removeCustomCellStyleHandler(predicate);
        return this;
    }
    public XlsExcelWriter addCustomCellStyleHandler(Predicate<Cell> predicate, BiFunction<Cell, CellStyleMo, CellStyleMo> handler) {
        this.sheetProperty.addCustomCellStyleHandler(predicate, handler);
        return this;
    }

    public XlsExcelWriter setFreeze(int freezeCol, int freezeRow) {
        this.sheetProperty.setFreeze(freezeCol, freezeRow);
        return this;
    }

    public XlsExcelWriter clearFreeze() {
        this.sheetProperty.setFreeze(0, 0);
        return this;
    }
}
