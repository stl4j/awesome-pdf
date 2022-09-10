package com.github.stl4j.awesomepdf.converter.excel;

import com.github.stl4j.awesomepdf.base.PdfCreator;
import com.github.stl4j.awesomepdf.converter.DocumentConverter;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfPCell;
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
     * The {@link PdfCreator} instance, for creating and rendering PDF document.
     */
    private final PdfCreator pdfCreator;

    /**
     * This {@link Set} will be used to specify which sheets of the Excel to be included in the PDF document.
     * All the sheets will be converted by default if target sheets not specified.
     */
    private final Set<Integer> targetSheets;

    /**
     * Default constructor for initializing something.
     */
    public ExcelConverter() {
        pdfCreator = new PdfCreator();
        targetSheets = new HashSet<>();
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
     * @param sheetIndex Index of the sheet number, zero-based, for example, 0, 1, 2, 3, etc.
     * @return Return {@code this} reference for method chain calls.
     * @see #sheets(int...)
     */
    public ExcelConverter sheet(int sheetIndex) {
        final int sheetCount = workbook.getNumberOfSheets();
        if (sheetIndex < 0 || sheetIndex > sheetCount - 1) {
            throw new IndexOutOfBoundsException(String.format("Sheet index %d out of range", sheetIndex));
        }
        targetSheets.add(sheetIndex);
        return this;
    }

    /**
     * This method is similar to {@link #sheet(int)}, the difference is that you can specify multiple parameters.
     *
     * @param sheetIndexes Index of the sheet number, zero-based, for example, 0, 1, 2, 3, etc.
     * @return Return {@code this} reference for method chain calls.
     * @see #sheet(int)
     */
    public ExcelConverter sheets(int... sheetIndexes) {
        for (final int sheetIndex : sheetIndexes) {
            sheet(sheetIndex);
        }
        return this;
    }

    /**
     * Resolve all the sheets, rows, and columns of the specified Excel file and convert them to a corresponding PDF document.
     *
     * @return Return {@code this} reference for method chain calls.
     * @throws DocumentException This exception may be thrown when opening the PDF document or adding components to it.
     * @see CellConverter#convert(Sheet, Cell)
     */
    public ExcelConverter convert() throws DocumentException {
        pdfCreator.newDocument();
        // Sheets
        final int sheetCount = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            // All sheets will be converted by default if no target sheet is specified.
            if (targetSheets.isEmpty() || targetSheets.contains(sheetIndex)) {
                pdfCreator.newPage();
                convertRowsAndColumns(sheetIndex);
            }
        }
        return this;
    }

    /**
     * Just for reducing the complexity of the method {@link #convert()}.
     *
     * @param sheetIndex Index of the sheet number, zero-based.
     * @throws DocumentException This exception may be thrown when adding components to the PDF document.
     * @see #convert()
     */
    private void convertRowsAndColumns(int sheetIndex) throws DocumentException {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        final int physicalCellCount = sheet.getRow(0).getPhysicalNumberOfCells();
        float[] pdfTableColumnWidths = new float[physicalCellCount];
        List<PdfPCell> pdfCells = new ArrayList<>();

        // Rows
        for (int rowIndex = 0, rowCount = sheet.getPhysicalNumberOfRows(); rowIndex < rowCount; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            // Cells
            for (int columnIndex = 0, cellCount = row.getPhysicalNumberOfCells(); columnIndex < cellCount; columnIndex++) {
                Cell cell = row.getCell(columnIndex);
                pdfTableColumnWidths[columnIndex] = 100;
                CellConverter cellConverter = new CellConverter();
                PdfPCell pdfCell = cellConverter.convert(sheet, cell);
                if (pdfCell != null) {
                    pdfCells.add(pdfCell);
                }
            }
        }

        pdfCreator.addTable(pdfTableColumnWidths, pdfCells);
    }

    /**
     * Save the PDF document to the specified path {@code targetFilePath} on the disk.
     *
     * @see DocumentConverter#save(String)
     */
    @Override
    public void save(String targetFilePath) throws IOException {
        pdfCreator.save(targetFilePath);
    }

}
