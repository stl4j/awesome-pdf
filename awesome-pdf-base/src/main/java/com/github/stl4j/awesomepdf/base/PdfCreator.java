package com.github.stl4j.awesomepdf.base;

import com.github.stl4j.awesomepdf.base.component.PdfTable;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

/**
 * This class is used to create a new PDF document, add the components
 * you want to render on it, and save the document as a PDF file.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @since 0.0.1
 */
public class PdfCreator {

    /**
     * The PDF document object of itext pdf library.
     */
    private Document document;

    /**
     * Output stream corresponding to PDF document content.
     */
    private ByteArrayOutputStream outputStream;

    /**
     * Create a new PDF document and open it.
     *
     * @throws DocumentException This exception may be thrown when creating or opening the document.
     */
    public void newDocument() throws DocumentException {
        document = new Document();
        outputStream = new ByteArrayOutputStream();
        PdfWriter writer = PdfWriter.getInstance(document, outputStream);
        writer.setPdfVersion(PdfWriter.PDF_VERSION_1_7);
        document.open();
    }

    /**
     * Create a new page for the PDF document.
     *
     * @see Document#newPage()
     */
    public void newPage() {
        document.newPage();
    }

    /**
     * Add a table component to the PDF document.
     *
     * @param columnWidths The column widths of the table component.
     * @param pdfCells     The cells you want to add to the table component.
     * @throws DocumentException This exception may be thrown when adding table component to the document.
     * @see PdfTable
     */
    public void addTable(float[] columnWidths, List<PdfPCell> pdfCells) throws DocumentException {
        PdfTable table = PdfTable.newTable(columnWidths).addCells(pdfCells);
        document.add(table.getTableElement());
    }

    /**
     * Save the PDF document as a file on the disk.
     *
     * @param targetFilePath The path of the PDF file to be saved.
     * @throws IOException This exception may be thrown when writing file to disk.
     */
    public void save(String targetFilePath) throws IOException {
        document.close();
        Files.write(Paths.get(targetFilePath), outputStream.toByteArray());
    }

}
