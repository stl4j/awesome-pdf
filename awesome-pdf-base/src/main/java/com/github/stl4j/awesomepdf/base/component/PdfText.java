package com.github.stl4j.awesomepdf.base.component;

import com.itextpdf.text.Chunk;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.ColumnText;

/**
 * This class is used to create a new text component for PDF document.
 *
 * @author stl4j - im.zhouchen@foxmail.com
 * @since 0.0.1
 */
public class PdfText {

    /**
     * The {@link ColumnText} element of itext pdf library.
     */
    private final ColumnText columnText;

    /**
     * The {@link Chunk} element of itext pdf library.
     */
    private final Chunk chunk;

    /**
     * Constructor with a string parameter to specify the text content of {@link #columnText}.
     *
     * @param text The text content of current component.
     */
    public PdfText(String text) {
        columnText = new ColumnText(null);
        chunk = text == null || text.trim().isEmpty() ? Chunk.createWhitespace(text) : new Chunk(text);
        columnText.addText(chunk);
    }

    /**
     * A shortcut method for creating a new {@link PdfText} object.
     */
    public static PdfText newText(String text) {
        return new PdfText(text);
    }

    /**
     * Set font for this text component.
     *
     * @param font A {@link Font} object of itext pdf library.
     * @return Return {@code this} reference for method chain calls.
     */
    public PdfText font(Font font) {
        chunk.setFont(font);
        return this;
    }

    /**
     * Return the {@link #columnText} reference.
     */
    public ColumnText build() {
        return columnText;
    }

}
