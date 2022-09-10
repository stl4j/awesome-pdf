package com.github.stl4j.awesomepdf.converter.excel;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;

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
            BaseColor pdfColor = convertExcelColor(excelCellColor);
            pdfCell.setBackgroundColor(pdfColor);
        }
    }

    /**
     * Convert Apache POI {@link Color} object to the {@link BaseColor} object of itext pdf library.
     *
     * @param color {@link Color} object of Apache POI library.
     * @return {@link BaseColor} object of itext pdf library.
     * @see HSSFColor
     * @see XSSFColor
     */
    private BaseColor convertExcelColor(Color color) {
        if (color instanceof HSSFColor) {
            HSSFColor hssfColor = HSSFColor.toHSSFColor(color);
            final short[] rgb = hssfColor.getTriplet();
            return new BaseColor(rgb[0], rgb[1], rgb[2]);
        } else if (color instanceof XSSFColor) {
            XSSFColor xssfColor = XSSFColor.toXSSFColor(color);
            final byte[] rgb = xssfColor.getRGB();
            // We need to convert the RGB value to [0, 255] from [-128, 127].
            return new BaseColor(rgb[0] & 0xff, rgb[1] & 0xff, rgb[2] & 0xff);
        }
        return null;
    }

}
