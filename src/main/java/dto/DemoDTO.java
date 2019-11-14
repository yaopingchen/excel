package dto;

import common.ExcelProperty;
import common.Style;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

public class DemoDTO {
    @ExcelProperty(index = 0,groupRow = "a")
    private String value1;
    @ExcelProperty(index = 1,groupRow = "a")
    private String value2;
    @ExcelProperty(index = 2)
    @Style(fontColor = IndexedColors.WHITE,backColor =IndexedColors.GREY_25_PERCENT,hAlignment = HorizontalAlignment.LEFT)
    private String value3;
    @ExcelProperty(index = 3,groupBy = "value3")
    @Style(bold = true)
    private String value4;

    public DemoDTO(String value1, String value2, String value3, String value4) {
        this.value1 = value1;
        this.value2 = value2;
        this.value3 = value3;
        this.value4 = value4;
    }

    public String getValue1() {
        return value1;
    }

    public void setValue1(String value1) {
        this.value1 = value1;
    }

    public String getValue2() {
        return value2;
    }

    public void setValue2(String value2) {
        this.value2 = value2;
    }

    public String getValue3() {
        return value3;
    }

    public void setValue3(String value3) {
        this.value3 = value3;
    }

    public String getValue4() {
        return value4;
    }

    public void setValue4(String value4) {
        this.value4 = value4;
    }
}
