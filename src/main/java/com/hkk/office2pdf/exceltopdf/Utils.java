package com.hkk.office2pdf.exceltopdf;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class Utils {

    public static byte[] getImage(Cell cell) {
        Sheet sheet = cell.getSheet();
        if (sheet instanceof HSSFSheet) {
            HSSFSheet hssfSheet = (HSSFSheet) sheet;
            if (hssfSheet.getDrawingPatriarch() != null) {
                List<HSSFShape> shapes = hssfSheet.getDrawingPatriarch().getChildren();
                for (HSSFShape shape : shapes) {
                    HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                    if (shape instanceof HSSFPicture) {
                        HSSFPicture pic = (HSSFPicture) shape;
                        PictureData data = pic.getPictureData();
                        int row1 = anchor.getRow1();
                        int col1 = anchor.getCol1();
                        if (row1 == cell.getRowIndex() && col1 == cell.getColumnIndex()) {
                            return data.getData();
                        }
                    }
                }
            }
        }
        return null;
    }

    public static int[] getColorRGB(Color color) {
        int red = 0;
        int green = 0;
        int blue = 0;
        int alpha = 0;

        if (color instanceof HSSFColor) {
            HSSFColor hssfColor = (HSSFColor) color;
            if (hssfColor.getIndex() == 27) {
                red = 218;
                green = 238;
                blue = 243;
            } else if (hssfColor.getIndex() == 31) {
                red = 197;
                green = 217;
                blue = 241;
            } else {
                short[] rgb = hssfColor.getTriplet();
                red = rgb[0];
                green = rgb[1];
                blue = rgb[2];
            }


        } else if (color instanceof XSSFColor) {
            XSSFColor xssfColor = (XSSFColor) color;
            byte[] rgb = xssfColor.getRgb();
            if (rgb != null) {
                red = (rgb[0] < 0) ? (rgb[0] + 256) : rgb[0];
                green = (rgb[1] < 0) ? (rgb[1] + 256) : rgb[1];
                blue = (rgb[2] < 0) ? (rgb[2] + 256) : rgb[2];
            }
        }

        if (red != 0 || green != 0 || blue != 0) {
            return new int[]{red, green, blue, alpha};
        } else return new int[]{255, 255, 255};
    }

    /**
     * 获取单元格颜色
     *
     * @param color 单元格颜色
     * @return RGB 三通道颜色, 默认: 黑色
     */
    public static int getRGB(Color color) {
        int result = 0x00FFFFFF;

        int red = 0;
        int green = 0;
        int blue = 0;

        if (color instanceof HSSFColor) {
            HSSFColor hssfColor = (HSSFColor) color;
            short[] rgb = hssfColor.getTriplet();
            red = rgb[0] < 0 ? rgb[0] + 255 : rgb[0];
            green = rgb[1] < 0 ? rgb[1] + 255 : rgb[1];
            blue = rgb[2] < 0 ? rgb[0] + 255 : rgb[2];
        }

        if (color instanceof XSSFColor) {
            XSSFColor xssfColor = (XSSFColor) color;
            byte[] rgb = xssfColor.getRgb();
            if (rgb != null) {
                red = (rgb[0] < 0) ? (rgb[0] + 256) : rgb[0];
                green = (rgb[1] < 0) ? (rgb[1] + 256) : rgb[1];
                blue = (rgb[2] < 0) ? (rgb[2] + 256) : rgb[2];
            }
        }

        if (red != 0 || green != 0 || blue != 0) {
            result = new java.awt.Color(red, green, blue).getRGB();
        }

        return result;
    }

    /**
     * 获取边框的颜色
     *
     * @param index 颜色版索引
     * @return RGB 三通道颜色
     */
    public static int getBorderRBG(Workbook workbook, short index) {
        int result = 0;

        if (workbook instanceof HSSFWorkbook) {
            HSSFWorkbook hwb = (HSSFWorkbook) workbook ;
            HSSFColor color = hwb.getCustomPalette().getColor(index);
            if (color != null) {
                result = getRGB(color);
            }
        }

        if (workbook  instanceof XSSFWorkbook) {
            XSSFColor color = new XSSFColor();
            color.setIndexed(index);
            result = getRGB(color);
        }

        return result;
    }
}
