package com.github.stl4j.awesomepdf.converter.support;

import java.io.OutputStream;

public interface Converter {

    void writeToPDF(OutputStream outputStream) throws Exception;

    void saveAsPDF(String targetFilePath) throws Exception;

    default String getDefaultTargetFilePath(String sourceFilePath) {
        if (sourceFilePath == null || sourceFilePath.trim().isEmpty()) {
            throw new IllegalArgumentException("'sourceFilePath' must not be null or empty");
        }
        final String targetFilePath = sourceFilePath.substring(0, sourceFilePath.lastIndexOf('.'));
        return targetFilePath + ".pdf";
    }

}
