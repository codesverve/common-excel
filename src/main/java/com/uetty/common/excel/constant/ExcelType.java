package com.uetty.common.excel.constant;

import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.IOException;
import java.io.InputStream;

// 暂时没用
@SuppressWarnings({"DeprecatedIsStillUsed", "unused"})
@Deprecated
public enum ExcelType {

    XLS(".xls"),
    XLSX(".xlsx");

    private String value;

    private ExcelType(String value) {
        this.setValue(value);
    }

    public String getValue() {
        return this.value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public static ExcelType valueOf(InputStream inputStream) {
        try {
            if (!inputStream.markSupported()) {
                return null;
            } else {
                FileMagic fileMagic = FileMagic.valueOf(inputStream);
                if (FileMagic.OLE2.equals(fileMagic)) {
                    return XLS;
                } else {
                    return FileMagic.OOXML.equals(fileMagic) ? XLSX : null;
                }
            }
        } catch (IOException var2) {
            throw new RuntimeException(var2);
        }
    }
}
