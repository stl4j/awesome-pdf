package com.github.stl4j.awesomepdf.converter.support;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashSet;
import java.util.Set;

/**
 * This class implements the interface {@link Converter}.
 * It will mainly responsible for converting an Excel file to the corresponding PDF document.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @see Converter
 * @since 0.0.1
 */
@SuppressWarnings("unused")
public class ExcelConverter implements Converter {

    /**
     * A {@link Workbook} object instance from the 3rd party library 'aspose-cells'.
     */
    private Workbook workbook;

    /**
     * We will use this if the caller does not specify the {@code targetFilePath}.
     * <p>
     * NOTICE: It will be ignored if we are reading Excel file from a java file stream.
     * In this case, the caller must specify the {@code targetFilePath}, otherwise an
     * exception {@link IllegalArgumentException} will be thrown.
     */
    private String defaultTargetFilePath;

    /**
     * This {@link Set} will be used to specify which sheets of the Excel to be included in the PDF document.
     * All the sheets will be visible by default if target sheets not specified.
     */
    private final Set<Integer> targetSheets;

    /**
     * Default constructor for initializing something.
     */
    public ExcelConverter() {
        targetSheets = new HashSet<>();
    }

    /**
     * Constructor that takes the {@code sourceFilePath} as an input parameter.
     *
     * @param sourceFilePath The source file path of the Excel document.
     * @throws Exception This exception will be thrown when reading the Excel document.
     */
    public ExcelConverter(String sourceFilePath) throws Exception {
        this();
        workbook = new Workbook(sourceFilePath);
        defaultTargetFilePath = getDefaultTargetFilePath(sourceFilePath);
    }

    /**
     * Constructor that takes the {@code inputStream} as an input parameter.
     *
     * @param inputStream The input stream of the Excel document content.
     * @throws Exception This exception will be thrown when reading the Excel document.
     */
    public ExcelConverter(InputStream inputStream) throws Exception {
        this();
        workbook = new Workbook(inputStream);
    }

    /**
     * Constructor that takes the {@code outputStream} as an input parameter.
     * <p>
     * This will be useful if you want to read the Excel from an output stream
     * into an input stream directly but you don't want to write its content to disk.
     *
     * @param outputStream The byte array output stream holding the Excel document content.
     * @throws Exception This exception will be thrown when reading the Excel document.
     */
    public ExcelConverter(ByteArrayOutputStream outputStream) throws Exception {
        this();
        ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());
        workbook = new Workbook(inputStream);
    }

    /**
     * Specify which sheet will be written to the PDF document.
     *
     * @param sheet The Excel sheet No. For example, 1, 2, 3, 4, etc.
     * @return Return {@code this} for method chaining.
     */
    public ExcelConverter sheet(int sheet) {
        targetSheets.add(sheet);
        return this;
    }

    /**
     * This method is similar to {@link #sheet(int)}, the difference is that you can specify multiple parameters.
     *
     * @param sheets The Excel sheet No. For example, 1, 2, 3, 4, etc.
     * @return Return {@code this} for method chaining.
     */
    public ExcelConverter sheets(int... sheets) {
        for (final int sheet : sheets) {
            targetSheets.add(sheet);
        }
        return this;
    }

    /**
     * @see Converter#writeToPdf(OutputStream)
     */
    @Override
    public void writeToPdf(OutputStream outputStream) throws Exception {
        processTargetSheetsVisibility(workbook);
        workbook.save(outputStream, getDefaultPdfSaveOptions());
    }

    /**
     * @see Converter#saveAsPdf(String)
     */
    @Override
    public void saveAsPdf(String targetFilePath) throws Exception {
        if (targetFilePath == null || targetFilePath.trim().isEmpty()) {
            if (defaultTargetFilePath == null) {
                throw new IllegalArgumentException("You must specify a target file path to save PDF if you are reading Excel from file stream");
            } else {
                targetFilePath = defaultTargetFilePath;
            }
        }
        processTargetSheetsVisibility(workbook);
        workbook.save(targetFilePath, getDefaultPdfSaveOptions());
    }

    /**
     * This method is similar to {@link #saveAsPdf(String)}, the difference is that you don't need to specify the {@code targetFilePath}.
     * In this case, the target file path will be equivalent to the {@code sourceFilePath}.
     *
     * @throws Exception This exception will be thrown when writing the PDF document.
     */
    public void saveAsPdf() throws Exception {
        saveAsPdf(null);
    }

    /**
     * This method is for processing which sheets of the Excel will be included in the PDF document when converting.
     *
     * @param workbook The {@link Workbook} object instance.
     */
    private void processTargetSheetsVisibility(Workbook workbook) {
        // All the sheets are visible by default if target sheets not specified.
        if (targetSheets.isEmpty()) {
            return;
        }

        WorksheetCollection worksheets = workbook.getWorksheets();
        for (int i = 0, size = worksheets.getCount(); i < size; i++) {
            Worksheet worksheet = worksheets.get(i);
            final int targetSheetNo = worksheet.getIndex() + 1;
            worksheet.setVisible(targetSheets.contains(targetSheetNo));
        }
    }

    /**
     * Return the default options of {@link PdfSaveOptions}.
     *
     * @return The default options of {@link PdfSaveOptions}.
     */
    private PdfSaveOptions getDefaultPdfSaveOptions() {
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setOnePagePerSheet(true);
        return saveOptions;
    }

}
