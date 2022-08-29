package com.github.stl4j.awesomepdf.converter.support;

import java.io.OutputStream;

/**
 * This interface is the root interface of all PDF converter classes.
 * A PDF converter can be used to convert some files to PDF documents.
 * For example, Excel, Word, HTML, SVG, etc. We need to implement this
 * interface according to the specific source file format.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see ExcelConverter
 * @since 0.0.1
 */
public interface Converter {

    /**
     * Write the PDF document content to the specified {@code outputStream}.
     *
     * @param outputStream The output stream to write the contents of a PDF document.
     * @throws Exception This exception will be thrown when writing the PDF document.
     */
    void writeToPdf(OutputStream outputStream) throws Exception;

    /**
     * Save the PDF document to the specified file path on the disk.
     *
     * @param targetFilePath The file path to save the PDF document.
     * @throws Exception This exception will be thrown when writing the PDF document.
     */
    void saveAsPdf(String targetFilePath) throws Exception;

    /**
     * Return a default file path for the PDF document to be saved based on the {@code sourceFilePath}.
     * This method will be useful if the caller does not specify the {@code targetFilePath}.
     * It will be equivalent to {@code sourceFilePath} by default.
     *
     * @param sourceFilePath The path to the source file to convert.
     * @return The path to the PDF document to be saved.
     */
    default String getDefaultTargetFilePath(String sourceFilePath) {
        if (sourceFilePath == null || sourceFilePath.trim().isEmpty()) {
            throw new IllegalArgumentException("'sourceFilePath' must not be null or empty");
        }
        final String targetFilePath = sourceFilePath.substring(0, sourceFilePath.lastIndexOf('.'));
        return targetFilePath + ".pdf";
    }

}
