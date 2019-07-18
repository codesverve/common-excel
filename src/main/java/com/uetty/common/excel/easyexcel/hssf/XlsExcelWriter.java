package com.uetty.common.excel.easyexcel.hssf;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.uetty.common.excel.easyexcel.hssf.handler.ModeledXlsWriterHandler;
import com.uetty.common.excel.model.CellStyleMo;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.function.BiFunction;
import java.util.function.Predicate;

@SuppressWarnings("unused")
public class XlsExcelWriter {

    private String outputPath;
    private Class<? extends BaseRowModel> modelClazz;
    private ModeledXlsWriterHandler modeledXlsWriterHandler;
    private boolean needHead = true;
    private int sheetIndex = 1;
    private String sheetName;

    public <T extends BaseRowModel> XlsExcelWriter(String outputPath, Class<T> modelClazz, int startRow) {
        this.outputPath = Objects.requireNonNull(outputPath);
        this.modelClazz = Objects.requireNonNull(modelClazz);
        modeledXlsWriterHandler = new ModeledXlsWriterHandler(modelClazz, startRow);
    }

    public void write(List<? extends BaseRowModel> list) throws IOException {
        if (!modeledXlsWriterHandler.hasExcelProperty()) { // 没有含标题名称的注解，就不需要标题了
            setNeedHead(false);
        }
        File file = new File(outputPath);
        try (FileOutputStream fos = new FileOutputStream(file)) {

            ExcelWriter writer = EasyExcelFactory.getWriterWithTempAndHandler(null, fos, ExcelTypeEnum.XLS, needHead, modeledXlsWriterHandler);

            Sheet sheet = new Sheet(sheetIndex, 0, modelClazz);
            int iStartRow = this.modeledXlsWriterHandler.getStartRow();
            if (modeledXlsWriterHandler.getHeadRowNum() == 0 || !needHead) { // easyexcel 本身在这一块有个小bug，兼容一下
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
    public ModeledXlsWriterHandler getModeledXlsWriterHandler() {
        return modeledXlsWriterHandler;
    }
    public void setModeledXlsWriterHandler(ModeledXlsWriterHandler modeledXlsWriterHandler) {
        this.modeledXlsWriterHandler = modeledXlsWriterHandler;
    }
    public boolean getNeedHead() {
        return needHead;
    }
    public XlsExcelWriter setNeedHead(boolean needHead) {
        this.needHead = needHead;
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
        this.modeledXlsWriterHandler.addMergeRange(firstRow, lastRow, firstCol, lastCol);
        return this;
    }
    public XlsExcelWriter removeMergeRange(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.modeledXlsWriterHandler.removeMergeRange(firstRow, lastRow, firstCol, lastCol);
        return this;
    }
    public XlsExcelWriter clearMergeRange() {
        this.modeledXlsWriterHandler.clearMergeRange();
        return this;
    }
    public XlsExcelWriter addExplicitConstraint(int firstRow, int lastRow, int firstCol, int lastCol, String[] explicitArray) {
        this.modeledXlsWriterHandler.addExplicitConstraint(firstRow, lastRow, firstCol, lastCol, explicitArray);
        return this;
    }
    public XlsExcelWriter removeRangeConstraint(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.modeledXlsWriterHandler.removeRangeConstraint(firstRow, lastRow, firstCol, lastCol);
        return this;
    }
    public XlsExcelWriter clearExplicitConstraints() {
        this.modeledXlsWriterHandler.clearExplicitConstraints();
        return this;
    }
    public XlsExcelWriter clearCustomCellStyleHandlers() {
        this.modeledXlsWriterHandler.clearCustomCellStyleHandlers();
        return this;
    }
    public XlsExcelWriter removeCustomCellStyleHandler(Predicate<Cell> predicate) {
        this.modeledXlsWriterHandler.removeCustomCellStyleHandler(predicate);
        return this;
    }
    public XlsExcelWriter addCustomCellStyleHandler(Predicate<Cell> predicate, BiFunction<Cell, CellStyleMo, CellStyleMo> handler) {
        this.modeledXlsWriterHandler.addCustomCellStyleHandler(predicate, handler);
        return this;
    }

    public XlsExcelWriter setFreeze(int freezeCol, int freezeRow) {
        this.modeledXlsWriterHandler.setFreeze(freezeCol,freezeRow);
        return this;
    }
}
