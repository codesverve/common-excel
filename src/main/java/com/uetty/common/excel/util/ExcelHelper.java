package com.uetty.common.excel.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Author: Vince
 * @Date: 2019/7/17 19:26
 */
public class ExcelHelper {

    private static SimpleDateFormat fullTimeFmt = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public static Object getCellValue(Cell cell, DateFormat readDateFormat) {
        Object value = null;
        if (cell != null) {
            CellType cellTypeEnum = cell.getCellTypeEnum();
            switch (cellTypeEnum) {
                case NUMERIC:
                    double numTxt = cell.getNumericCellValue();
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = new Date(0);
                        date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                        DateFormat sdf = fullTimeFmt;
                        if (readDateFormat != null) {
                            sdf = readDateFormat;
                        }
                        value = sdf.format(date);

                    } else {
                        // 只能这样处理下整型了
                        if (numTxt == ((long) numTxt)) {
                            value = String.valueOf((long) numTxt);
                        } else {
                            value = String.valueOf(numTxt);
                        }

                    }// 全部当做文本
                    break;
                case BOOLEAN:
                    value = cell.getBooleanCellValue();
                    break;
                case BLANK:
                    value = null;
                    break;
                case FORMULA:
                    FormulaEvaluator eval;
                    Workbook workbook = cell.getSheet().getWorkbook();
                    if (workbook instanceof HSSFWorkbook) {
                        eval = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
                    } else {
                        eval = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
                    }
                    eval.evaluateInCell(cell);
                    value = getCellValue(cell, readDateFormat);
                    break;
                case STRING:
                    RichTextString rtxt = cell.getRichStringCellValue();
                    if (rtxt == null) {
                        break;
                    }
                    // 全角空格转为半角空格
                    value = rtxt.getString().replace("　", " ");
                    break;
                default:
            }
        }
        return value;
    }
}
