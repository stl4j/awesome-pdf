package com.github.stl4j.awesomepdf.converter.excel;

import com.itextpdf.text.BaseColor;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.List;

/**
 * Various utility functions that make working with a cells and rows easier.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @since 0.0.1
 */
public class CellConverterUtils {

    private CellConverterUtils() {
        // An utility class should not be instantiated by external classes.
    }

    /**
     * Calculates the row span and column span, return an array of int values.
     * The first element is the row span, the second is the column span.
     *
     * @param excelSheet {@link Sheet} object of Apache POI library.
     * @param excelCell  {@link Cell} object of Apache POI library.
     * @return An array of int values, the first element is the row span, the second is the column span.
     */
    public static int[] calculateRowAndColumnSpan(Sheet excelSheet, Cell excelCell) {
        final int[] rowAndColumnSpan = new int[]{1, 1};
        final int rowIndex = excelCell.getRowIndex();
        final int columnIndex = excelCell.getColumnIndex();
        List<CellRangeAddress> mergedRegions = excelSheet.getMergedRegions();
        for (CellRangeAddress mergedRegion : mergedRegions) {
            final int firstRow = mergedRegion.getFirstRow();
            final int lastRow = mergedRegion.getLastRow();
            final int firstColumn = mergedRegion.getFirstColumn();
            final int lastColumn = mergedRegion.getLastColumn();
            if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                final int rowSpan = lastRow - firstRow + 1;
                final int columnSpan = lastColumn - firstColumn + 1;
                rowAndColumnSpan[0] = rowSpan;
                rowAndColumnSpan[1] = columnSpan;
            }
        }
        return rowAndColumnSpan;
    }

    /**
     * Check that the cell is a merged cell.
     *
     * @param excelSheet {@link Sheet} object of Apache POI library.
     * @param excelCell  {@link Cell} object of Apache POI library.
     * @return {@code true} if the cell is a merged cell, {@code false} otherwise.
     */
    public static boolean isMergedCell(Sheet excelSheet, Cell excelCell) {
        final int[] rowAndColumnSpan = calculateRowAndColumnSpan(excelSheet, excelCell);
        return rowAndColumnSpan[0] > 1 || rowAndColumnSpan[1] > 1;
    }

    /**
     * Convert Apache POI {@link Color} object to the {@link BaseColor} object of itext pdf library.
     *
     * @param color {@link Color} object of Apache POI library.
     * @return {@link BaseColor} object of itext pdf library.
     * @see HSSFColor
     * @see XSSFColor
     */
    public static BaseColor convertExcelColor(Color color) {
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
