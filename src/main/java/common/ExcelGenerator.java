package common;

import dto.ExcelCell;
import dto.ExcelCellStyle;
import dto.Section;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class ExcelGenerator {
    public Workbook buildExcel(String sheetName, List<Section> sections){
        HSSFWorkbook workbook = new HSSFWorkbook();
        addSheet(workbook,sheetName,sections);
        return workbook;
    }
    public Sheet addSheet(Workbook workbook, String sheetName, List<Section> sections){
        Sheet sheet = workbook.createSheet(sheetName);
        int startRow = 0;
        for(Section section:sections){
            startRow = addSectionToSheet(sheet,section,startRow);
        }
        return sheet;
    }
    private CellStyle getStyle(Workbook workbook, ExcelCellStyle excelCellStyle){
        CellStyle style = getDefaultStyle(workbook);
        if(excelCellStyle != null){
            if(excelCellStyle.gethAlignment()!=null){
                style.setAlignment(excelCellStyle.gethAlignment());
            }
            if(excelCellStyle.getvAlignment()!=null){
                style.setVerticalAlignment(excelCellStyle.getvAlignment());
            }
            if(excelCellStyle.getWrapText()!=null){
                style.setWrapText(excelCellStyle.getWrapText());
            }
            if(StringUtils.isNotBlank(excelCellStyle.getFontName())||
                    excelCellStyle.getFontSize()!=0||
                    excelCellStyle.getFontColor()!=null){
                Font font=workbook.createFont();
                if(StringUtils.isNotBlank(excelCellStyle.getFontName())){
                    font.setFontName(excelCellStyle.getFontName());
                }else {
                    font.setFontName("MS PGothic");
                }
                if(excelCellStyle.getFontSize()!=0){
                    font.setFontHeightInPoints(excelCellStyle.getFontSize());
                }else {
                    font.setFontHeightInPoints((short) 10);
                }
                if(excelCellStyle.getFontColor()!=null){
                    font.setColor(excelCellStyle.getFontColor().getIndex());
                }
                if (excelCellStyle.getBold()){
                    font.setBold(excelCellStyle.getBold());
                }
                style.setFont(font);
            }
            if(StringUtils.isNotBlank(excelCellStyle.getDataFormat())){
                DataFormat dataFormat = workbook.createDataFormat();
                style.setDataFormat(dataFormat.getFormat(excelCellStyle.getDataFormat()));
            }
            if(excelCellStyle.getBackColor()!=null){
                style.setFillForegroundColor(excelCellStyle.getBackColor().getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if(excelCellStyle.getBorder()!=null){
                style.setBorderBottom(excelCellStyle.getBorder());
                style.setBorderTop(excelCellStyle.getBorder());
                style.setBorderLeft(excelCellStyle.getBorder());
                style.setBorderRight(excelCellStyle.getBorder());
            }
        }
        return style;
    }
    public CellStyle getDefaultStyle(Workbook workbook){
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        Font font=workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("MS PGothic");
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
    private void mergedRegion(Sheet sheet,int startRow, int endRow, int startCol, int endCol){
        if(startRow==endRow&&startCol==endCol){
            return;
        }
        CellRangeAddress region = new CellRangeAddress(startRow, endRow, startCol, endCol);
        sheet.addMergedRegion(region);
    }
    class GroupRecord{
        private int startRow;
        private int startCol;
        private String group;

        public int getStartRow() {
            return startRow;
        }

        public void setStartRow(int startRow) {
            this.startRow = startRow;
        }

        public String getGroup() {
            return group;
        }

        public void setGroup(String group) {
            this.group = group;
        }

        public int getStartCol() {
            return startCol;
        }

        public void setStartCol(int startCol) {
            this.startCol = startCol;
        }
    }


    /**
     * 添加section到sheet中
     * @param sheet
     * @param section
     * @param currentRow section开始行
     * @return 结束行行号
     */
    private int addSectionToSheet(Sheet sheet, Section section,int currentRow) {
        if(section.getOffsetRow()!=0){
            currentRow = currentRow + section.getOffsetRow();
        }
        int startCol = section.getOffsetCol();
        int colCount = section.getData()[0].length;
        //加标题
        currentRow = addSectionData(sheet,section.getTitles(),currentRow,startCol);
        //加数据
        currentRow = addSectionData(sheet,section.getData(),currentRow,startCol);

        //加foot 汇总行
        return addFooter(sheet,section.getFooter(),currentRow,startCol,colCount);
    }
    private int addFooter(Sheet sheet, ExcelCell[] data, int startRow,int startCol,int colCount){
        if(data==null){
            return startRow;
        }
        int offsetCol = colCount-data.length+startCol;
        Row row = sheet.createRow(startRow);
        addRow(sheet, data, offsetCol, row);
        return ++startRow;
    }

    private void addRow(Sheet sheet, ExcelCell[] data, int offsetCol, Row row) {
        for(int j=0;j<data.length;j++){
            ExcelCell myCell = data[j];
            if(myCell==null){
                continue;
            }
            Cell cell = row.createCell(offsetCol+j);
            cell.setCellValue(myCell.getValue());
            if(myCell.getStyle()!=null){
                cell.setCellStyle(getStyle(sheet.getWorkbook(),myCell.getStyle()));
            }else {
                cell.setCellStyle(getDefaultStyle(sheet.getWorkbook()));
            }
        }
    }

    private int addSectionData(Sheet sheet, ExcelCell[][] data, int startRow, int startCol) {
        if (data==null||data.length==0){
            return startRow;
        }
       for(int i = 0; i<data.length; i++){
           ExcelCell[] rowData = data[i];
           Row row = sheet.createRow(i+startRow);
           addRow(sheet, rowData, startCol, row);
       }
       mergeData(sheet,data,startRow,startCol);
       return data.length+startRow;
    }

    /**
     * 合并
     * @param sheet
     * @param data
     * @param startRow
     * @param startCol
     */
    private void mergeData(Sheet sheet, ExcelCell[][] data, int startRow, int startCol){
        //横向合并
        for(int i = 0; i<data.length; i++){
            ExcelCell[] rowData = data[i];
            GroupRecord currentGroup = null;
            for(int j=0;j<rowData.length;j++){
                ExcelCell myCell = rowData[j];
                if(currentGroup!=null){
                    if(!currentGroup.getGroup().equals(myCell.getGroup())){
                        mergedRegion(sheet,i+startRow,i+startRow,currentGroup.getStartCol()+startCol,j-1+startCol);
                        if(myCell.getGroup()!=null){
                            currentGroup.setGroup(myCell.getGroup());
                            currentGroup.setStartCol(j);
                        }else {
                            currentGroup = null;
                        }
                    }
                }else {
                    if(myCell.getGroup()!=null){
                        currentGroup = new GroupRecord();
                        currentGroup.setGroup(myCell.getGroup());
                        currentGroup.setStartCol(j);
                    }else {
                        currentGroup = null;
                    }
                }
            }
            if(currentGroup!=null){
                mergedRegion(sheet,i+startRow,i+startRow,currentGroup.getStartCol()+startCol,rowData.length-1+startCol);
            }
        }
        //纵向合并
        for(int j=0;j<data[0].length;j++){
            GroupRecord currentGroup = null;
            for(int i = 0; i<data.length; i++){
                ExcelCell myCell = data[i][j];
                if(currentGroup!=null){
                    if(!currentGroup.getGroup().equals(myCell.getGroup())){
                        mergedRegion(sheet,currentGroup.getStartRow()+startRow,i-1+startRow,j+startCol,j+startCol);
                        if(myCell.getGroup()!=null){
                            currentGroup.setGroup(myCell.getGroup());
                            currentGroup.setStartRow(i);
                        }else {
                            currentGroup = null;
                        }
                    }
                }else {
                    if(myCell.getGroup()!=null){
                        currentGroup = new GroupRecord();
                        currentGroup.setGroup(myCell.getGroup());
                        currentGroup.setStartRow(i);
                    }else {
                        currentGroup = null;
                    }
                }

            }
            if(currentGroup!=null){
                mergedRegion(sheet,currentGroup.getStartRow()+startRow,data.length-1+startRow,j+startCol,j+startCol);
            }
        }

    }
}
