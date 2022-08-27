package com.github.stl4j.awesomepdf.converter.support;

public interface ExcelConverter extends Converter {
    void saveAsPDF(String targetFilePath, int[] targetSheets) throws Exception;
}
