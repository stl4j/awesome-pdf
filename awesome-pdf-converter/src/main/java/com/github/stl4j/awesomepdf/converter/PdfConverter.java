package com.github.stl4j.awesomepdf.converter;

import com.github.stl4j.awesomepdf.converter.excel.ExcelConverter;

import java.io.IOException;
import java.io.InputStream;

/**
 * This class is an utility class, which provides some static methods
 * to make it easier to use this library, you don't need to instantiate it.
 * <p>
 * Just use like this:
 * <blockquote><pre>
 *     PdfConverter.fromExcel("test.xlsx").saveAsPdf("test.pdf");
 * </pre></blockquote>
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see ExcelConverter
 * @since 0.0.1
 */
@SuppressWarnings("unused")
public final class PdfConverter {

    private PdfConverter() {
        // An utility class should not be instantiated by external classes.
    }

    /**
     * @param sourceFilePath The file path of the Excel document.
     * @return The {@link ExcelConverter} instance, for easy to use chain calls of method.
     * @throws IOException This exception will be thrown when reading the Excel document.
     * @see ExcelConverter
     */
    public static ExcelConverter fromExcel(String sourceFilePath) throws IOException {
        return new ExcelConverter(sourceFilePath);
    }

    /**
     * @param inputStream The input stream of the Excel document content.
     * @return The {@link ExcelConverter} instance, for easy to use chain calls of method.
     * @throws IOException This exception will be thrown when reading the Excel document.
     * @see ExcelConverter
     */
    public static ExcelConverter fromExcel(InputStream inputStream) throws IOException {
        return new ExcelConverter(inputStream);
    }

}
