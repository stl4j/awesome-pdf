package com.github.stl4j.awesomepdf.converter.excel;

import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.ss.usermodel.*;

import java.util.EnumMap;

/**
 * This class is used to convert the properties about cell dimensions,
 * such as width, height, row span, column span, alignment, etc.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see CellAdapter
 * @since 0.0.1
 */
public class CellDimensionAdapter implements CellAdapter {

    /**
     * This Map is used to hold the mapping between Apache POI {@link HorizontalAlignment} and itext pdf {@link Element} ALIGN_XXX.
     */
    private static final EnumMap<HorizontalAlignment, Integer> HORIZONTAL_ALIGNMENT_MAP = new EnumMap<>(HorizontalAlignment.class);

    /**
     * This Map is used to hold the mapping between Apache POI {@link VerticalAlignment} and itext pdf {@link Element} ALIGN_XXX.
     */
    private static final EnumMap<VerticalAlignment, Integer> VERTICAL_ALIGNMENT_MAP = new EnumMap<>(VerticalAlignment.class);

    /**
     * Default constructor, initialize some default values.
     */
    public CellDimensionAdapter() {
        HORIZONTAL_ALIGNMENT_MAP.put(HorizontalAlignment.LEFT, Element.ALIGN_LEFT);
        HORIZONTAL_ALIGNMENT_MAP.put(HorizontalAlignment.GENERAL, Element.ALIGN_LEFT);
        HORIZONTAL_ALIGNMENT_MAP.put(HorizontalAlignment.CENTER, Element.ALIGN_CENTER);
        HORIZONTAL_ALIGNMENT_MAP.put(HorizontalAlignment.RIGHT, Element.ALIGN_RIGHT);
        VERTICAL_ALIGNMENT_MAP.put(VerticalAlignment.TOP, Element.ALIGN_TOP);
        VERTICAL_ALIGNMENT_MAP.put(VerticalAlignment.CENTER, Element.ALIGN_MIDDLE);
        VERTICAL_ALIGNMENT_MAP.put(VerticalAlignment.BOTTOM, Element.ALIGN_BOTTOM);
    }

    @Override
    public void adapt(Sheet excelSheet, Cell excelCell, PdfPCell pdfCell) {
        adaptRowAndColumnSpan(excelSheet, excelCell, pdfCell);
        adaptAlignment(excelCell, pdfCell);
    }

    /**
     * Adapt row span and column span properties.
     *
     * @param excelSheet {@link Sheet} object of Apache POI library.
     * @param excelCell  {@link Cell} object of Apache POI library.
     * @param pdfCell    {@link PdfPCell} object of itext pdf library.
     */
    private void adaptRowAndColumnSpan(Sheet excelSheet, Cell excelCell, PdfPCell pdfCell) {
        final int[] rowAndColumnSpan = CellConverterUtils.calculateRowAndColumnSpan(excelSheet, excelCell);
        pdfCell.setRowspan(rowAndColumnSpan[0]);
        pdfCell.setColspan(rowAndColumnSpan[1]);
    }

    /**
     * Adapt alignment properties, including horizontal alignment and vertical alignment.
     *
     * @param excelCell {@link Cell} object of Apache POI library.
     * @param pdfCell   {@link PdfPCell} object of itext pdf library.
     */
    private void adaptAlignment(Cell excelCell, PdfPCell pdfCell) {
        CellStyle cellStyle = excelCell.getCellStyle();
        final HorizontalAlignment horizontalAlignment = cellStyle.getAlignment();
        final VerticalAlignment verticalAlignment = cellStyle.getVerticalAlignment();
        pdfCell.setHorizontalAlignment(HORIZONTAL_ALIGNMENT_MAP.get(horizontalAlignment));
        pdfCell.setVerticalAlignment(VERTICAL_ALIGNMENT_MAP.get(verticalAlignment));
        pdfCell.setUseAscender(true);
    }

}
