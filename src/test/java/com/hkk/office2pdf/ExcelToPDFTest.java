package com.hkk.office2pdf;

import com.hkk.office2pdf.exceltopdf.ExcelConverter;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.PageSize;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.IOException;
import java.net.URL;

@SpringBootTest
public class ExcelToPDFTest {

    @Test
    public void excel2003toPDF() throws DocumentException, InvalidFormatException, IOException {

        String path = "OpinionNotice.xls";
        URL url = getClass().getClassLoader().getResource(path);
        ExcelConverter converter = new ExcelConverter(PageSize.A4.rotate());
        converter.toPdf(url.getFile());
    }

    @Test
    public void excel2007toPDF() throws DocumentException, InvalidFormatException, IOException {

        String path = "OpinionNotice.xlsx";
        URL url = getClass().getClassLoader().getResource(path);
        ExcelConverter converter = new ExcelConverter(PageSize.A4.rotate());
        converter.toPdf(url.getFile());
    }

}
