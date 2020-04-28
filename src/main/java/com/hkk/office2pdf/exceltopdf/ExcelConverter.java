package com.hkk.office2pdf.exceltopdf;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import java.io.*;

public class ExcelConverter {

    private Excel excel;

    private PdfDocument document;

    /**
     * 默认纸张为A4
     */
    public ExcelConverter() {
        this.document = new PdfDocument();
    }

    /**
     * 设置纸张大小
     *
     * @param pageSize
     */
    public ExcelConverter(Rectangle pageSize) {

        this.excel = new Excel();
        this.document = new PdfDocument(pageSize);
    }

    public byte[] toPdf(String path) throws IOException, DocumentException, InvalidFormatException {
        File file = new File(path);
        if (!file.exists()) {
            // 文件不存在
            throw new RuntimeException("文件不存在，请联系管理员");
        }
        FileInputStream inputStream = new FileInputStream(file);

        return toPdf(inputStream);
    }

    /**
     * 转
     *
     * @param stream
     * @return
     * @throws IOException
     * @throws DocumentException
     * @throws InvalidFormatException
     */
    public byte[] toPdf(InputStream stream) throws IOException, DocumentException, InvalidFormatException {
        this.excel.prepare(stream);
        this.document.prepare();

        for (Row row : this.excel.getSheet()) {
            for (Cell cell : row) {
                if (this.excel.mergedCell(cell)) {
                    continue;
                }

                PdfPCell pdfpCell = getPdfCell(row, cell);
                this.document.addCell(pdfpCell);
            }
        }

        document.finish(this.excel.getRelativeWidths());

        return null;
    }

    /**
     * 获取PDFCell
     *
     * @param row
     * @param cell
     * @return
     * @throws BadElementException
     * @throws IOException
     */
    private PdfPCell getPdfCell(Row row, Cell cell) throws DocumentException, IOException {
        CellStyle cellStyle = cell.getCellStyle();

        cell.setCellType(Cell.CELL_TYPE_STRING);
        PdfPCell pdfpCell = new PdfPCell();

        //单元格背景色 TODO这个背景色不准，需要优化调色版
        Color color = cellStyle.getFillForegroundColorColor();
        BaseColor baseColor = this.document.getBaseColor(color);
        pdfpCell.setBackgroundColor(baseColor);

        //如果是被合并的单元格
        CellRangeAddress range = this.excel.getCellRange(cell);
        if (range != null) {
            //单元格占列数
            pdfpCell.setColspan(range.getLastColumn() - range.getFirstColumn() + 1);
            //单元格占行数
            pdfpCell.setRowspan(range.getLastRow() - range.getFirstRow() + 1);
        }

        //单元格垂直对齐
        pdfpCell.setVerticalAlignment(getVerticalAlign(cellStyle.getVerticalAlignment()));
        //单元格水平对齐
        pdfpCell.setHorizontalAlignment(getHorizonAlign(cellStyle.getAlignment()));
        //单元格内容
        pdfpCell.setPhrase(this.document.getPhrase(cell,getFont(cellStyle)));

        if (this.excel.getSheet().getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
            pdfpCell.setFixedHeight(this.getPixelHeight(row.getHeightInPoints()));
        }

        //单元格边框
        addBorder(pdfpCell, cellStyle);
        //单元格图片
        addImage(pdfpCell, cell);
        return pdfpCell;
    }

    private void addImage(PdfPCell pdfpCell, Cell cell) throws BadElementException, IOException {
        byte[] bytes = Utils.getImage(cell);
        if (bytes != null) {
            pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            Image image = Image.getInstance(bytes);
            pdfpCell.setImage(image);
        }
    }

    private float getPixelHeight(float poiHeight) {
        float pixel = poiHeight / 28.6f * 26f;
        return pixel;
    }

    private Font getFont(CellStyle style) {
        int fontColorIndex = 0;
        Workbook workbook = excel.getWorkbook();
        Color fontColor = null;
        if (style instanceof HSSFCellStyle) {
            HSSFCellStyle hssfCellStyle = (HSSFCellStyle) style;
            HSSFFont font = hssfCellStyle.getFont(workbook);
            fontColorIndex = font.getColor();
            fontColor = font.getHSSFColor((HSSFWorkbook) workbook);
        } else {
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle) style;
            XSSFFont font = xssfCellStyle.getFont();
            fontColorIndex = font.getColor();
            fontColor = font.getXSSFColor();
        }
        // 字体样式索引
        short index = style.getFontIndex();
        org.apache.poi.ss.usermodel.Font font = workbook.getFontAt(index);
        Font result = Resource.getFont(font);

        //是否加粗
        if (org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD == font.getBoldweight()) {
            result.setStyle(Font.BOLD);
        }

        //字体颜色
        if (fontColor != null && fontColorIndex != 8) {
            BaseColor baseColor = this.document.getBaseColor(fontColor);
            result.setColor(baseColor);
        }

        //下划线
        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if (FontUnderline.SINGLE == underline) {
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }

    protected void addBorder(PdfPCell cell , CellStyle style) {
        Workbook wb = excel.getWorkbook();
        cell.setBorderColorLeft(new BaseColor(Utils.getBorderRBG(wb,style.getLeftBorderColor())));
        cell.setBorderColorRight(new BaseColor(Utils.getBorderRBG(wb,style.getRightBorderColor())));
        cell.setBorderColorTop(new BaseColor(Utils.getBorderRBG(wb,style.getTopBorderColor())));
        cell.setBorderColorBottom(new BaseColor(Utils.getBorderRBG(wb,style.getBottomBorderColor())));
    }

    private void addBorder1(PdfPCell cell, CellStyle style) {

        Workbook workbook = excel.getWorkbook();
        short borderBottom = style.getBorderBottom();
        short borderLeft = style.getBorderLeft();
        short borderRight = style.getBorderRight();
        short borderTop = style.getBorderTop();

        if (borderLeft > 0) {
            HSSFColor hssfColorByIndex = ExcelColor.getHssfColorByIndex(style.getLeftBorderColor());
            short[] rgb = hssfColorByIndex.getTriplet();
            cell.setBorderWidthLeft(borderLeft);
            cell.setBorderColorLeft(new BaseColor(rgb[0],rgb[1],rgb[2]));
        } else {
            cell.disableBorderSide(PdfPCell.LEFT);
        }

        if (borderRight > 0) {
            cell.setBorderWidthRight(borderRight);
            cell.setBorderColorRight(new BaseColor(Utils.getBorderRBG(workbook, style.getRightBorderColor())));
        } else {
            cell.disableBorderSide(PdfPCell.RIGHT);
        }

        if (borderTop > 0) {
            cell.setBorderWidthTop(borderTop);
            cell.setBorderColorTop(new BaseColor(Utils.getBorderRBG(workbook, style.getTopBorderColor())));
        } else {
            cell.disableBorderSide(PdfPCell.TOP);
        }

        if (borderBottom > 0) {
            cell.setBorderWidthBottom(borderBottom);
            cell.setBorderColorBottom(new BaseColor(Utils.getBorderRBG(workbook, style.getBottomBorderColor())));
        } else {
            cell.disableBorderSide(PdfPCell.BOTTOM);
        }
    }

    private int getVerticalAlign(short align) {
        int result = 0;
        switch (align) {
            case CellStyle.VERTICAL_BOTTOM:
                result = Element.ALIGN_BOTTOM;
                break;
            case CellStyle.VERTICAL_CENTER:
                result = Element.ALIGN_MIDDLE;
                break;
            case CellStyle.VERTICAL_JUSTIFY:
                result = Element.ALIGN_JUSTIFIED;
                break;
            case CellStyle.VERTICAL_TOP:
                result = Element.ALIGN_TOP;
                break;
        }
        return result;
    }

    private int getHorizonAlign(short align) {
        int result = 0;
        switch (align) {
            case CellStyle.ALIGN_LEFT:
                result = Element.ALIGN_LEFT;
                break;
            case CellStyle.ALIGN_RIGHT:
                result = Element.ALIGN_RIGHT;
                break;
            case CellStyle.ALIGN_JUSTIFY:
                result = Element.ALIGN_JUSTIFIED;
                break;
            case CellStyle.ALIGN_CENTER:
                result = Element.ALIGN_CENTER;
                break;
        }
        return result;
    }
}