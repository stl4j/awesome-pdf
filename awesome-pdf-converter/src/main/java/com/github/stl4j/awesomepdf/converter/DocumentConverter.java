package com.github.stl4j.awesomepdf.converter;

import com.github.stl4j.awesomepdf.converter.excel.ExcelConverter;
import com.itextpdf.text.DocumentException;

import java.io.IOException;

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
public interface DocumentConverter {
    /**
     * Save the PDF document to the specified file path on the disk.
     *
     * @param targetFilePath The file path to save the PDF document.
     * @throws DocumentException This exception may be thrown when writing the PDF document.
     * @throws IOException       This exception may be thrown when writing the PDF document.
     */
    void save(String targetFilePath) throws DocumentException, IOException;
}
