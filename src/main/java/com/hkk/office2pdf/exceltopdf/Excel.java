package com.hkk.office2pdf.exceltopdf;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

public class Excel {

    private Workbook workbook;

    private Sheet sheet;

    public Excel() {

    }

    public Excel(InputStream stream) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(stream);
        this.sheet = this.workbook.getSheetAt(0);
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public void prepare(InputStream stream) throws IOException, InvalidFormatException {
        this.workbook = WorkbookFactory.create(stream);
        this.sheet = this.workbook.getSheetAt(0);
    }

    public float[] getRelativeWidths() {
        Row row = this.sheet.getRow(0);
        short columns = row.getLastCellNum();

        float[] cw = new float[columns];
        for (int i = 0; i < columns; i++) {
            cw[i] = this.sheet.getColumnWidth(i);
        }

        return cw;
    }

    public CellRangeAddress getCellRange(Cell cell) {
        CellRangeAddress result = null;

        int num = this.sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = this.sheet.getMergedRegion(i);
            if (range.getFirstColumn() == cell.getColumnIndex() && range.getFirstRow() == cell.getRowIndex()) {
                result = range;
            }
        }
        return result;
    }

    public boolean mergedCell(Cell cell) {
        int num = this.sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = this.sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            if (cell.getRowIndex() == firstRow && cell.getColumnIndex() == firstColumn) {
                return false;
            }
            if (firstRow <= cell.getRowIndex() && lastRow >= cell.getRowIndex()) {
                if (firstColumn <= cell.getColumnIndex() && lastColumn >= cell.getColumnIndex()) {
                    return true;
                }
            }
        }

        return false;
    }


}
