package test;

import common.ExcelGenerator;
import common.ExcelHelper;
import dto.DemoDTO;
import dto.ExcelCell;
import dto.ExcelCellStyle;
import dto.Section;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class test {
    public static void main(String[] args) {
        ExcelGenerator excelGenerator = new ExcelGenerator();
        List<Section> sections = mockData();
        Workbook excel = excelGenerator.buildExcel("test",sections);
//        excelGenerator.addSheet(excel,"xxx",sections);
        File file = new File("C:\\Users\\yaopchen\\Desktop\\test.xls");
        try {
            excel.write(new FileOutputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<Section> mockData() {
        List<Section> sections = new ArrayList<Section>();
        sections.add(mockSection());
        Section section2 = mockSection();
        section2.setOffsetRow(1);
        section2.setOffsetCol(1);
        LinkedHashMap<String,List> map = new LinkedHashMap<String, List>();
        map.put("C",Arrays.asList("5","6","7"));
        map.put("B",Arrays.asList("8","9","10"));
        ExcelHelper.addDynamicColumn(section2,map,3);
        sections.add(section2);
        return sections;
    }

    private static ExcelCell[] mockFooter() {
        ExcelCell[] footer = new ExcelCell[2];
        ExcelCellStyle excelCellStyle = new ExcelCellStyle();
        excelCellStyle.setWrapText(false);
        footer[0]=new ExcelCell("今期の源泉徴収すべき予納税額の内、現地より予納した税額",excelCellStyle);
        footer[1]=new ExcelCell("20");
        return footer;
    }

    private static Section mockSection() {
        Section section = new Section();
        List<DemoDTO> list = new ArrayList<DemoDTO>();
        list.add(new DemoDTO("1","2","1","4"));
        list.add(new DemoDTO("2","2","3","4"));
        list.add(new DemoDTO("3","2","3","4"));
        ExcelCell[][] data = ExcelHelper.convertData(list);
        section.setData(data);
        ExcelCellStyle style = new ExcelCellStyle();
        style.setBackColor(IndexedColors.LIGHT_TURQUOISE);
        section.setTitles(ExcelHelper.getSimpleTitles(Arrays.asList("title1","title2","title3","title4"),style));
        section.setFooter(mockFooter());
        return section;
    }

}
