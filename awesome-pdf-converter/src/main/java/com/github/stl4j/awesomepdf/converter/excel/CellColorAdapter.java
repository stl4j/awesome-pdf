package com.github.stl4j.awesomepdf.converter.excel;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * This class is used to convert the properties about cell colors,
 * such as background color, font color, border color, etc.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see CellAdapter
 * @since 0.0.1
 */
public class CellColorAdapter implements CellAdapter {

    @Override
    public void adapt(Sheet excelSheet, Cell excelCell, PdfPCell pdfCell) {
        adaptBackgroundColor(excelCell.getCellStyle(), pdfCell);
    }

    /**
     * Adapt background color property.
     *
     * @param excelCellStyle {@link CellStyle} object of Apache POI library.
     * @param pdfCell        {@link PdfPCell} object of itext pdf library.
     */
    private void adaptBackgroundColor(CellStyle excelCellStyle, PdfPCell pdfCell) {
        Color excelCellColor = excelCellStyle.getFillForegroundColorColor();
        if (excelCellColor != null) {
            BaseColor pdfColor = CellConverterUtils.convertExcelColor(excelCellColor);
            pdfCell.setBackgroundColor(pdfColor);
        }
    }

}
