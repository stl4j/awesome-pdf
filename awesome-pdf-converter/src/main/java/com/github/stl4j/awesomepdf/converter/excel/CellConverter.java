package com.github.stl4j.awesomepdf.converter.excel;

import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * This class is intended to convert a {@link Cell} object to a {@link PdfPCell} object.
 * It will hold all the implementations of the {@link CellAdapter} interface and call their methods
 * {@link CellAdapter#adapt(Sheet, Cell, PdfPCell)}, finally a {@link PdfPCell} object would be returned.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see CellAdapter
 * @since 0.0.1
 */
public class CellConverter {

    /**
     * All the implementations of the {@link CellConverter} interface.
     */
    private final List<CellAdapter> cellAdapters;

    /**
     * Default constructor, initialize something.
     */
    public CellConverter() {
        cellAdapters = new ArrayList<>();
        cellAdapters.add(new CellValueAdapter());
        cellAdapters.add(new CellColorAdapter());
        cellAdapters.add(new CellDimensionAdapter());
    }

    /**
     * Convert the given Excel {@link Cell} object to the corresponding {@link PdfPCell} object.
     *
     * @param excelSheet {@link Sheet} object of Apache POI library.
     * @param excelCell  {@link Cell} object of Apache POI library.
     * @return {@link PdfPCell} object of itext pdf library.
     * @see ExcelConverter#convert()
     */
    public PdfPCell convert(Sheet excelSheet, Cell excelCell) {
        final String stringCellValue = excelCell.getStringCellValue();
        // When the string cell value is null or empty, maybe this cell has a merged region.
        // We only need the non-null and non-empty value to create new PdfPCell object in this case.
        if (stringCellValue != null && stringCellValue.trim().length() > 0) {
            PdfPCell pdfCell = new PdfPCell();
            for (CellAdapter cellAdapter : cellAdapters) {
                cellAdapter.adapt(excelSheet, excelCell, pdfCell);
            }
            return pdfCell;
        }
        return null;
    }

}
