package dto;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class ExcelCellStyle {
    private String fontName;
    private short fontSize;
    private HorizontalAlignment hAlignment;
    private VerticalAlignment vAlignment;
    private IndexedColors fontColor;
    private IndexedColors backColor;
    private String dataFormat;
    private Boolean wrapText;
    private Boolean bold;
    private BorderStyle border;

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public short getFontSize() {
        return fontSize;
    }

    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    public HorizontalAlignment gethAlignment() {
        return hAlignment;
    }

    public void sethAlignment(HorizontalAlignment hAlignment) {
        this.hAlignment = hAlignment;
    }

    public VerticalAlignment getvAlignment() {
        return vAlignment;
    }

    public void setvAlignment(VerticalAlignment vAlignment) {
        this.vAlignment = vAlignment;
    }

    public IndexedColors getFontColor() {
        return fontColor;
    }

    public void setFontColor(IndexedColors fontColor) {
        this.fontColor = fontColor;
    }

    public IndexedColors getBackColor() {
        return backColor;
    }

    public void setBackColor(IndexedColors backColor) {
        this.backColor = backColor;
    }

    public String getDataFormat() {
        return dataFormat;
    }

    public void setDataFormat(String dataFormat) {
        this.dataFormat = dataFormat;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    public Boolean getBold() {
        return bold;
    }

    public void setBold(Boolean bold) {
        this.bold = bold;
    }

    public BorderStyle getBorder() {
        return border;
    }

    public void setBorder(BorderStyle border) {
        this.border = border;
    }
}
