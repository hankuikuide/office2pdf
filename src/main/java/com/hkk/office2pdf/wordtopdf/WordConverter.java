package com.hkk.office2pdf.wordtopdf;


import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class WordConverter {
    public void toPdf(String docPath, String pdfPath) throws IOException {
        InputStream stream = new FileInputStream(new File(docPath));
        XWPFDocument document = new XWPFDocument(stream);
        PdfOptions options = PdfOptions.create();
        OutputStream outputStream = new FileOutputStream(new File(pdfPath));
        PdfConverter.getInstance().convert(document, outputStream, options);
    }
}
