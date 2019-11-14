package dto;

public class ExcelCell {
    private String value;
    private String group;
    private ExcelCellStyle style;

    public ExcelCell(String value, String group, ExcelCellStyle style) {
        this.value = value;
        this.group = group;
        this.style = style;
    }
    public ExcelCell(String value, ExcelCellStyle style) {
        this.value = value;
        this.style = style;
    }
    public ExcelCell(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getGroup() {
        return group;
    }

    public void setGroup(String group) {
        this.group = group;
    }

    public ExcelCellStyle getStyle() {
        return style;
    }

    public void setStyle(ExcelCellStyle style) {
        this.style = style;
    }
}
