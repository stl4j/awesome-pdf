package com.github.stl4j.awesomepdf.base.component;

import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;

import java.util.List;

/**
 * This class is used to create a new table component for PDF document.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @since 0.0.1
 */
public class PdfTable {

    /**
     * The {@link PdfPTable} element of itext pdf library.
     */
    private final PdfPTable tableElement;

    /**
     * Constructor with a {@code float} array parameter to specify the column widths of {@link #tableElement}.
     *
     * @param columnWidths The column widths of {@link #tableElement}.
     */
    public PdfTable(float[] columnWidths) {
        tableElement = new PdfPTable(columnWidths);
        tableElement.setWidthPercentage(100);
    }

    /**
     * A shortcut method for creating a new {@link PdfTable} object.
     */
    public static PdfTable newTable(float[] columnWidths) {
        return new PdfTable(columnWidths);
    }

    /**
     * Add cells for the {@link #tableElement}.
     *
     * @param cells The cells to be added to the table element.
     */
    public PdfTable addCells(List<PdfPCell> cells) {
        for (PdfPCell cell : cells) {
            tableElement.addCell(cell);
        }
        return this;
    }

    /**
     * Return the {@link #tableElement} reference.
     */
    public PdfPTable getTableElement() {
        return tableElement;
    }

}
