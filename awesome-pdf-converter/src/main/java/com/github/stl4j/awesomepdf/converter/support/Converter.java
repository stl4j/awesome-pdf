package com.github.stl4j.awesomepdf.converter.support;

public interface Converter {

    default String guessTargetFilePath(String sourceFilePath) {
        if (sourceFilePath == null || sourceFilePath.trim().isEmpty()) {
            throw new IllegalArgumentException("'sourceFilePath' must not be null or empty");
        }
        final String targetFilePath = sourceFilePath.substring(0, sourceFilePath.lastIndexOf('.'));
        return targetFilePath + ".pdf";
    }

}
