package common;

import dto.ExcelCell;
import dto.ExcelCellStyle;
import dto.Section;

import java.lang.reflect.Field;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.UUID;

public class ExcelHelper {
    public static ExcelCell[][] getSimpleTitles(List<String> list,ExcelCellStyle style){
        int length = list.size();
        ExcelCell[][] titles = new ExcelCell[1][length];
        for(int i=0;i<length;i++){
            titles[0][i]=new ExcelCell(list.get(i),style);
        }
        return titles;
    }
    public static ExcelCell[][] convertData(List<?> list){
        Class clazz = list.get(0).getClass();
        Field[] fields = clazz.getDeclaredFields();
        int length = fields.length;
        int rowCount = list.size();
        ExcelCell[][] data = new ExcelCell[rowCount][length];
        for(int i =0;i<rowCount;i++){
            Object obj = list.get(i);
            String suffix = UUID.randomUUID().toString();
            for(Field field :fields){
                if (field.isAnnotationPresent(ExcelProperty.class)) {
                    field.setAccessible(true);
                    ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
                    Style style = field.getAnnotation(Style.class);
                    ExcelCellStyle excelCellStyle = null;
                    if(style!=null){
                        excelCellStyle = getStyle(style);
                    }
                    int index = excelProperty.index();
                    String group = getGroup(excelProperty,clazz,obj,suffix,index);
                    try {
                        String value = field.get(obj).toString();
                        data[i][index] = new ExcelCell(value,group,excelCellStyle);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        return data;
    }

    private static ExcelCellStyle getStyle(Style style) {
        ExcelCellStyle excelCellStyle = new ExcelCellStyle();
        excelCellStyle.setFontName(style.fontName());
        excelCellStyle.setFontSize(style.fontSize());
        excelCellStyle.setFontColor(style.fontColor());
        excelCellStyle.setBackColor(style.backColor());
        excelCellStyle.sethAlignment(style.hAlignment());
        excelCellStyle.setvAlignment(style.vAlignment());
        excelCellStyle.setDataFormat(style.dataFormat());
        excelCellStyle.setBold(style.bold());
        excelCellStyle.setBorder(style.border());
        return excelCellStyle;
    }

    private static String getGroup(ExcelProperty excelProperty, Class clazz, Object obj,String suffix,int index){
        String group = null;
        if(StringUtils.isNotBlank(excelProperty.groupRow())){
            group = excelProperty.groupRow()+suffix;
        }else if(StringUtils.isNotBlank(excelProperty.groupBy())){
            Field field = null;
            try {
                field = clazz.getDeclaredField(excelProperty.groupBy());
                field.setAccessible(true);
                return field.get(obj).toString()+index;
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        return group;
    }

    /**
     * 动态添加列(仅支持单行标题)
     * @param section
     * @param map
     * @param index
     */
    public static void addDynamicColumn(Section section, LinkedHashMap<String,List> map, int index){
        int oldLength = section.getTitles()[0].length;
        int addLength = map.keySet().size();
        int newLength = addLength+oldLength;
        int dataRowCount = section.getData().length;
        int titleRowCount = section.getTitles().length;
        int oldDataLength = section.getData()[0].length;
        int newDataLength = addLength+oldDataLength;
        ExcelCell[][] titles= new ExcelCell[titleRowCount][newLength];
        ExcelCell[][] data= new ExcelCell[dataRowCount][newDataLength];
        //标题扩容
        for(int i=0;i<section.getTitles().length;i++){
            System.arraycopy(section.getTitles()[i],0,titles[i],0,index);
            System.arraycopy(section.getTitles()[i],index,titles[i],index+addLength,oldLength-index);
        }
        //data扩容
        for(int i=0;i<section.getData().length;i++){
            System.arraycopy(section.getData()[i],0,data[i],0,index);
            System.arraycopy(section.getData()[i],index,data[i],index+addLength,oldDataLength-index);
        }
        int styleBaseCol = 0;
        //复制前一列的样式
        if(index>0){
            styleBaseCol = index - 1;
        }
        for(String title:map.keySet()){
            List<String> values = map.get(title);
            if(values==null){
                continue;
            }
            titles[0][index]= new ExcelCell(title,titles[0][styleBaseCol].getStyle());
            for(int row=0;row<values.size();row++){
                data[row][index] = new ExcelCell(values.get(row), data[row][styleBaseCol].getStyle());
            }
            index++;
        }
        section.setData(data);
        section.setTitles(titles);
    }
}