package com.github.stl4j.awesomepdf.converter.support;

import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

import java.util.HashSet;
import java.util.Set;

public class ExcelConverter implements Converter {

    private final String sourceFilePath;

    private final Set<Integer> targetSheets;

    public ExcelConverter(String sourceFilePath) {
        this.sourceFilePath = sourceFilePath;
        targetSheets = new HashSet<>();
    }

    public ExcelConverter sheet(int sheet) {
        targetSheets.add(sheet);
        return this;
    }

    public ExcelConverter sheets(int... sheets) {
        for (final int sheet : sheets) {
            targetSheets.add(sheet);
        }
        return this;
    }

    @Override
    public void saveAsPDF(String targetFilePath) throws Exception {
        targetFilePath = targetFilePath == null || targetFilePath.trim().isEmpty() ? guessTargetFilePath(sourceFilePath) : targetFilePath;
        Workbook workbook = new Workbook(sourceFilePath);
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setOnePagePerSheet(true);
        processTargetSheetsVisibility(workbook);
        workbook.save(targetFilePath);
    }

    public void saveAsPDF() throws Exception {
        saveAsPDF(null);
    }

    private void processTargetSheetsVisibility(Workbook workbook) {
        // All sheets are visible by default if target sheets not specified
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

}
