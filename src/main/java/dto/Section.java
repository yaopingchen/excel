package dto;

public class Section {
    private ExcelCell[][] titles;
    private ExcelCell[][] data;
    private ExcelCell[] footer;
    /**
     * 偏移行数
     */
    private int offsetRow;
    /**
     * 偏移列数
     */
    private int offsetCol;

    public ExcelCell[][] getTitles() {
        return titles;
    }

    public void setTitles(ExcelCell[][] titles) {
        this.titles = titles;
    }

    public ExcelCell[][] getData() {
        return data;
    }

    public void setData(ExcelCell[][] data) {
        this.data = data;
    }

    public ExcelCell[] getFooter() {
        return footer;
    }

    public void setFooter(ExcelCell[] footer) {
        this.footer = footer;
    }

    public int getOffsetRow() {
        return offsetRow;
    }

    public void setOffsetRow(int offsetRow) {
        this.offsetRow = offsetRow;
    }

    public int getOffsetCol() {
        return offsetCol;
    }

    public void setOffsetCol(int offsetCol) {
        this.offsetCol = offsetCol;
    }
}
