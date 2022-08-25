package com.github.stl4j.awesomepdf.converter;

import com.github.stl4j.awesomepdf.converter.support.DefaultExcelConverter;

public abstract class PdfConverter {

    public static DefaultExcelConverter fromExcel(String sourceFilePath) {
        return new DefaultExcelConverter(sourceFilePath);
    }

}
