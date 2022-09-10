package com.github.stl4j.awesomepdf.converter.excel;

import com.github.stl4j.awesomepdf.base.component.PdfText;
import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * This class is used to convert the properties about cell value.
 * In general, the cell value will be a simple string, also it could
 * be some special values, such as image, chart, phone number, money amount, etc.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see CellAdapter
 * @since 0.0.1
 */
public class CellValueAdapter implements CellAdapter {

    @Override
    public void adapt(Sheet excelSheet, Cell excelCell, PdfPCell pdfCell) {
        adaptStringValue(excelCell, pdfCell);
    }

    /**
     * Adapt the simple string value.
     *
     * @param excelCell {@link Cell} object of Apache POI library.
     * @param pdfCell   {@link PdfPCell} object of itext pdf library.
     */
    private void adaptStringValue(Cell excelCell, PdfPCell pdfCell) {
        final String cellValue = excelCell.getStringCellValue();
        pdfCell.setColumn(PdfText.newText(cellValue).build());
    }

}
