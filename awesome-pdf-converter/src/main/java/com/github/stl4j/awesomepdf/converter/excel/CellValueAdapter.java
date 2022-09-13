package com.github.stl4j.awesomepdf.converter.excel;

import com.github.stl4j.awesomepdf.base.component.PdfText;
import com.github.stl4j.awesomepdf.base.util.FontUtils;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfPCell;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

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

    /**
     * The {@link Workbook} reference.
     */
    private final Workbook workbook;

    /**
     * Default constructor for initializing something.
     *
     * @param workbook The {@link Workbook} reference.
     */
    public CellValueAdapter(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void adapt(Sheet excelSheet, Cell excelCell, PdfPCell pdfCell) throws DocumentException, IOException {
        adaptStringValue(excelCell, pdfCell);
    }

    /**
     * Adapt the simple string value.
     *
     * @param excelCell {@link Cell} object of Apache POI library.
     * @param pdfCell   {@link PdfPCell} object of itext pdf library.
     */
    private void adaptStringValue(Cell excelCell, PdfPCell pdfCell) throws DocumentException, IOException {
        final String cellValue = excelCell.getStringCellValue();
        com.itextpdf.text.Font cellFont = adaptCellFont(excelCell);
        pdfCell.setColumn(PdfText.newText(cellValue).font(cellFont).build());
    }

    /**
     * Adapt the font for this cell value.
     *
     * @param excelCell {@link Cell} object of Apache POI library.
     * @return {@link com.itextpdf.text.Font} object of itext pdf library.
     * @throws DocumentException This exception may be thrown when creating the {@link BaseFont} object.
     * @throws IOException       This exception may be thrown when creating the {@link BaseFont} object.
     */
    private com.itextpdf.text.Font adaptCellFont(Cell excelCell) throws DocumentException, IOException {
        CellStyle cellStyle = excelCell.getCellStyle();
        final int fontIndex = cellStyle.getFontIndex();
        org.apache.poi.ss.usermodel.Font excelFont = workbook.getFontAt(fontIndex);
        final String fontPath = FontUtils.findSystemFontPath(excelFont.getFontName());
        BaseFont pdfBaseFont = BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED, BaseFont.CACHED);
        final short fontSize = excelFont.getFontHeightInPoints();
        final int fontStyleValue = adaptFontStyleValue(excelFont);
        BaseColor fontColor = adaptFontColor(excelFont);
        return new com.itextpdf.text.Font(pdfBaseFont, fontSize, fontStyleValue, fontColor);
    }

    /**
     * Adapt the font style value from {@link org.apache.poi.ss.usermodel.Font} to {@link com.itextpdf.text.Font}.
     *
     * @param excelFont {@link org.apache.poi.ss.usermodel.Font} object of Apache POI library.
     * @return The font style value from {@link com.itextpdf.text.Font}.
     */
    private int adaptFontStyleValue(org.apache.poi.ss.usermodel.Font excelFont) {
        StringBuilder fontStyle = new StringBuilder(com.itextpdf.text.Font.FontStyle.NORMAL.getValue());
        if (excelFont.getBold()) {
            fontStyle.append(com.itextpdf.text.Font.FontStyle.BOLD.getValue());
        }
        if (excelFont.getItalic()) {
            fontStyle.append(com.itextpdf.text.Font.FontStyle.ITALIC.getValue());
        }
        if (excelFont.getUnderline() == org.apache.poi.ss.usermodel.Font.U_SINGLE) {
            fontStyle.append(com.itextpdf.text.Font.FontStyle.UNDERLINE.getValue());
        }
        return com.itextpdf.text.Font.getStyleValue(fontStyle.toString());
    }

    /**
     * Adapt the font color from {@link HSSFColor} or {@link XSSFColor} to {@link BaseColor}.
     *
     * @param excelFont {@link org.apache.poi.ss.usermodel.Font} object of Apache POI library.
     * @return {@link BaseColor} object of itext pdf library.
     */
    private BaseColor adaptFontColor(org.apache.poi.ss.usermodel.Font excelFont) {
        if (workbook instanceof HSSFWorkbook && excelFont instanceof HSSFFont) {
            HSSFFont hssfFont = (HSSFFont) excelFont;
            HSSFColor hssfColor = hssfFont.getHSSFColor((HSSFWorkbook) workbook);
            return CellConverterUtils.convertExcelColor(hssfColor);
        } else if (workbook instanceof XSSFWorkbook && excelFont instanceof XSSFFont) {
            XSSFFont xssfFont = (XSSFFont) excelFont;
            XSSFColor xssfColor = xssfFont.getXSSFColor();
            return CellConverterUtils.convertExcelColor(xssfColor);
        }
        return BaseColor.BLACK;
    }

}
