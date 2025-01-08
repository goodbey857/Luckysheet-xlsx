package top.kingdon.converter;

import com.alibaba.fastjson.JSONArray;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.kingdon.converter.sheet.SheetConverter;
import top.kingdon.parser.Parser;
import top.kingdon.parser.sheet.LuckySheet;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;

public class Converter {
    public static void convert(){
        // TODO 转sheet
        // TODO 转workBook
    }

    public static void main(String[] args) {
        try{
            Workbook workbook = new XSSFWorkbook();
//            Workbook workbook = new SXSSFWorkbook();      // 这种方式inlineStyle样式不能正常渲染
            SheetConverter sheetConverter = new SheetConverter(workbook);
            InputStream resourceAsStream = Parser.class.getClassLoader().getResourceAsStream("demoSheetData.json");
            byte[] bytes = new byte[resourceAsStream.available()];
            resourceAsStream.read(bytes);
            String jsonString = new String(bytes);
            JSONArray sheetsJson = JSONArray.parseArray(jsonString);
            for (int i = 0; i < sheetsJson.size(); i++) {
                if(Arrays.asList(4,7).contains(i))continue; // Sparkline, pivotTable暂时不能导出
                sheetConverter.convert(Parser.parser(sheetsJson.getString(i)));
            }

            FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
            workbook.write(fileOut);

        }catch (Exception e){
            e.printStackTrace();
        }



    }
}
