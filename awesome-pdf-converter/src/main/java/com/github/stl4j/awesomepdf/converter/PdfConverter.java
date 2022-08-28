package com.github.stl4j.awesomepdf.converter;

import com.github.stl4j.awesomepdf.converter.support.ExcelConverter;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

/**
 * This class is an utility class, which provides some static methods
 * to make it easier to use this library, you don't need to instantiate it.
 * <p>
 * Just use like this:
 * <blockquote><pre>
 *     PdfConverter.fromExcel("test.xslx").saveAsPdf();
 *     PdfConverter.fromExcel("test.xslx").saveAsPdf("test.pdf");
 * </pre></blockquote><p>
 *
 * @author stl4j <im.zhouchen@foxmail.com>
 * @see ExcelConverter
 * @since 0.0.1
 */
@SuppressWarnings("unused")
public final class PdfConverter {

    private PdfConverter() {
        // An utility class should not be instantiated directly.
    }

    /**
     * @see ExcelConverter(String)
     */
    public static ExcelConverter fromExcel(String sourceFilePath) throws Exception {
        return new ExcelConverter(sourceFilePath);
    }

    /**
     * @see ExcelConverter(InputStream)
     */
    public static ExcelConverter fromExcel(InputStream inputStream) throws Exception {
        return new ExcelConverter(inputStream);
    }

    /**
     * @see ExcelConverter(ByteArrayOutputStream)
     */
    public static ExcelConverter fromExcel(ByteArrayOutputStream outputStream) throws Exception {
        return new ExcelConverter(outputStream);
    }

}
