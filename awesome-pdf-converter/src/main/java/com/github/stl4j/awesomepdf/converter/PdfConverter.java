package com.github.stl4j.awesomepdf.converter;

import com.github.stl4j.awesomepdf.converter.support.ExcelConverter;

import java.io.InputStream;

public abstract class PdfConverter {

    public static ExcelConverter fromExcel(String sourceFilePath) throws Exception {
        return new ExcelConverter(sourceFilePath);
    }

    public static ExcelConverter fromExcel(InputStream inputStream) throws Exception {
        return new ExcelConverter(inputStream);
    }

}
