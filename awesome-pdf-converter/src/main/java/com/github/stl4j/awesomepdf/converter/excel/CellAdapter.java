package com.github.stl4j.awesomepdf.converter.excel;

import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * This interface is the root interface of all Excel cell adapter classes.
 * A cell adapter can be used to convert Apache POI Excel cell styles or
 * properties to itext pdf table cell properties, for example, width, height,
 * border, padding, alignment, font, background, etc. We need to implement this
 * interface according to the different cell properties.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see CellColorAdapter
 * @see CellDimensionAdapter
 * @see CellValueAdapter
 * @since 0.0.1
 */
public interface CellAdapter {
    /**
     * Adapt the given {@code excelCell} of {@code excelSheet} to a {@code pdfCell}.
     *
     * @param excelSheet {@link Sheet} object of Apache POI library.
     * @param excelCell  {@link Cell} object of Apache POI library.
     * @param pdfCell    {@link PdfPCell} object of itext pdf library.
     */
    void adapt(Sheet excelSheet, Cell excelCell, PdfPCell pdfCell);
}
