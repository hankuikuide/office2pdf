package com.hkk.office2pdf;

import com.hkk.office2pdf.exceltopdf.ExcelConverter;
import com.hkk.office2pdf.wordtopdf.WordConverter;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.IOException;
import java.net.URL;

@SpringBootTest
public class WordToPDFTest {
    @Test
    public void word2007toPDF() throws DocumentException, InvalidFormatException, IOException {

        String path = "OpinionNotice.docx";
        URL url = getClass().getClassLoader().getResource(path);
        WordConverter converter = new WordConverter();
        converter.toPdf(url.getFile(),"OpinionNotice.pdf");
    }
}
