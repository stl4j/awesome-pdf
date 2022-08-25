package com.github.stl4j.awesomepdf.converter.support;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;

public class DefaultExcelConverter implements ExcelConverter {

    private static final String AUTO_GUESS_TARGET_FILE_PATH = null;

    private static final int[] DEFAULT_FIRST_SHEET = null;

    private final String sourceFilePath;

    public DefaultExcelConverter(String sourceFilePath) {
        this.sourceFilePath = sourceFilePath;
    }

    @Override
    public void toPDF(String targetFilePath, int[] targetSheets) throws Exception {
        targetFilePath = targetFilePath == null || targetFilePath.trim().isEmpty() ? guessTargetFilePath(sourceFilePath) : targetFilePath;
        Workbook workbook = new Workbook(sourceFilePath);
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setOnePagePerSheet(true);
        processTargetSheetsVisibility(workbook, targetSheets);
        workbook.save(targetFilePath);
    }

    public void toPDF(int[] targetSheets) throws Exception {
        toPDF(AUTO_GUESS_TARGET_FILE_PATH, targetSheets);
    }

    public void toPDF() throws Exception {
        toPDF(AUTO_GUESS_TARGET_FILE_PATH, DEFAULT_FIRST_SHEET);
    }

    private void processTargetSheetsVisibility(Workbook workbook, int[] targetSheets) {
        WorksheetCollection worksheets = workbook.getWorksheets();
        if (targetSheets == null || targetSheets.length == 0) {
            worksheets.get(0).setVisible(true);
        } else {
            for (int i = 0; i < targetSheets.length; i++) {
                worksheets.get(i).setVisible(true);
            }
        }
    }

}
