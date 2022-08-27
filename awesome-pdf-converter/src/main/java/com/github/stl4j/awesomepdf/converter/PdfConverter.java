package com.github.stl4j.awesomepdf.converter;

import com.github.stl4j.awesomepdf.converter.support.ExcelConverter;

public abstract class PdfConverter {

    public static ExcelConverter fromExcel(String sourceFilePath) {
        return new ExcelConverter(sourceFilePath);
    }

}
