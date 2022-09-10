package com.github.stl4j.awesomepdf.converter.excel;

import com.github.stl4j.awesomepdf.base.PdfCreator;
import com.github.stl4j.awesomepdf.converter.DocumentConverter;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * This class implements the interface {@link DocumentConverter}.
 * It will mainly responsible for converting an Excel file to the corresponding PDF document.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see DocumentConverter
 * @since 0.0.1
 */
@SuppressWarnings("unused")
public class ExcelConverter implements DocumentConverter {

    /**
     * The {@link Workbook} instance of Apache POI library.
     */
    private Workbook workbook;

    /**
     * This {@link Set} will be used to specify which sheets of the Excel to be included in the PDF document.
     * All the sheets will be converted by default if target sheets not specified.
     */
    private final Set<Integer> targetSheets;

    /**
     * This array holds the width value of each column in the PDF {@link PdfPTable} component.
     */
    private float[] pdfTableColumnWidths;

    /**
     * This list will contain all the target {@link PdfPCell} instance.
     */
    private final List<PdfPCell> pdfCells;

    /**
     * Default constructor for initializing something.
     */
    public ExcelConverter() {
        targetSheets = new HashSet<>();
        pdfCells = new ArrayList<>();
    }

    /**
     * Constructor that takes the {@code sourceFilePath} as an input parameter.
     *
     * @param sourceFilePath The source file path of the Excel document.
     * @throws IOException This exception may be thrown when reading the Excel document.
     */
    public ExcelConverter(String sourceFilePath) throws IOException {
        this();
        workbook = WorkbookFactory.create(new File(sourceFilePath));
    }

    /**
     * Constructor that takes the {@code inputStream} as an input parameter.
     *
     * @param inputStream The input stream of the Excel document content.
     * @throws IOException This exception may be thrown when reading the Excel document.
     */
    public ExcelConverter(InputStream inputStream) throws IOException {
        this();
        workbook = WorkbookFactory.create(inputStream);
    }

    /**
     * Specify which sheet will be written to the PDF document.
     *
     * @param sheet Index of the sheet number, zero-based, for example, 0, 1, 2, 3, etc.
     * @return Return {@code this} reference for method chain calls.
     * @see #sheets(int...)
     */
    public ExcelConverter sheet(int sheet) {
        targetSheets.add(sheet);
        return this;
    }

    /**
     * This method is similar to {@link #sheet(int)}, the difference is that you can specify multiple parameters.
     *
     * @param sheets Index of the sheet number, zero-based, for example, 0, 1, 2, 3, etc.
     * @return Return {@code this} reference for method chain calls.
     * @see #sheet(int)
     */
    public ExcelConverter sheets(int... sheets) {
        for (final int sheet : sheets) {
            targetSheets.add(sheet);
        }
        return this;
    }

    /**
     * Resolve all the sheets, rows, and columns of the specified Excel file and convert them to a corresponding PDF document.
     *
     * @return Return {@code this} reference for method chain calls.
     * @see CellConverter#convert(Sheet, Cell)
     */
    public ExcelConverter convert() {
        // Sheets
        final int sheetCount = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            // All sheets will be converted by default if no target sheet is specified.
            if (targetSheets.isEmpty() || targetSheets.contains(sheetIndex)) {
                convertRowsAndColumns(sheetIndex);
            }
        }
        return this;
    }

    /**
     * Just for reducing the complexity of the method {@link #convert()}.
     *
     * @param sheetIndex Index of the sheet number, zero-based.
     * @see #convert()
     */
    private void convertRowsAndColumns(int sheetIndex) {
        // Rows
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        final int rowCount = sheet.getPhysicalNumberOfRows();
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            // Cells
            Row row = sheet.getRow(rowIndex);
            final int cellCount = row.getPhysicalNumberOfCells();
            pdfTableColumnWidths = new float[cellCount];
            for (int columnIndex = 0; columnIndex < cellCount; columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                pdfTableColumnWidths[columnIndex] = 100;
                CellConverter cellConverter = new CellConverter();
                PdfPCell pdfCell = cellConverter.convert(sheet, cell);
                if (pdfCell != null) {
                    pdfCells.add(pdfCell);
                }
            }
        }
    }

    /**
     * Save the PDF document to the specified path {@code targetFilePath} on the disk.
     *
     * @see DocumentConverter#save(String)
     */
    @Override
    public void save(String targetFilePath) throws DocumentException, IOException {
        PdfCreator creator = new PdfCreator();
        creator.newDocument().addTable(pdfTableColumnWidths, pdfCells).save("test.pdf");
    }

}
