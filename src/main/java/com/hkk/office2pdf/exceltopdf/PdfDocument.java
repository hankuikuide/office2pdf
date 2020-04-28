package com.hkk.office2pdf.exceltopdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Color;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PdfDocument {

    private Document document;

    private List<PdfPCell> pdfPCells = new ArrayList<>();

    protected boolean setting = false;

    public PdfDocument() {
        this.document = new Document();
    }
    public PdfDocument(Rectangle pageSize) {
        this.document = new Document(pageSize);
    }

    public void prepare() throws IOException, DocumentException {
        //PdfWriter实例
        PdfWriter.getInstance(this.document, new FileOutputStream("filename.pdf"));
        // 打开文档
        this.document.open();
    }

    public void finish(float[] cw) throws DocumentException {

        PdfPTable pdfPTable = new PdfPTable(cw);
        pdfPTable.setWidthPercentage(100);

        pdfPCells.forEach(c -> pdfPTable.addCell(c));

        this.document.add(pdfPTable);

        this.document.close();
    }

    public void addCell(PdfPCell cell) {
        this.pdfPCells.add(cell);
    }


    public Phrase getPhrase(Cell cell,Font font) {
        if (this.setting ) {
            return new Phrase(cell.getStringCellValue(), font);
        }
        Anchor anchor = new Anchor(cell.getStringCellValue(), font);
        this.setting = true;
        return anchor;
    }

    public BaseColor getBaseColor(Color color) {
        int[] colorRGB = Utils.getColorRGB(color);
        return new BaseColor(colorRGB[0], colorRGB[1], colorRGB[2], 0xff);

    }

}
