package top.kingdon;

import com.alibaba.fastjson2.JSON;
import org.apache.poi.ss.usermodel.Workbook;
import top.kingdon.converter.sheet.SheetConverter;
import top.kingdon.parser.Parser;
import top.kingdon.utils.Util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class LuckysheetXlsx {

    public static Workbook export(String sheetsJson, int exportIndex) {
        SheetConverter sheetConverter = new SheetConverter();
        if (JSON.isValidArray(sheetsJson)) {
            List<String> jsonStringList = Util.splitJsonArray(sheetsJson);
            if(exportIndex != -1){
                sheetConverter.convert(Parser.parser(jsonStringList.get(exportIndex)));
            }else{
                for (String jsonString : jsonStringList) {
                    sheetConverter.convert(Parser.parser(jsonString));
                }
            }

        }else{
            sheetConverter.convert(Parser.parser(sheetsJson));
        }

        return sheetConverter.getWorkbook();
    }

    public static void export(String sheetsJson,String order, OutputStream outputStream) {
        int exportIndex = -1;
        if(order != null && order != "" && order != "all"){
            exportIndex = Integer.parseInt(order);
        }
        Workbook workbook = export(sheetsJson, exportIndex);
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void exportToLocalFile(String sheetsJson, String filePath) {
        if(!filePath.endsWith(".xlsx")){
            filePath = filePath + ".xlsx";
        }
        try {
            FileOutputStream fileOut = new FileOutputStream(filePath);
            Workbook export = export(sheetsJson, -1);
            export.write(fileOut);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String[] args) {

        try{
            InputStream resourceAsStream = Parser.class.getClassLoader().getResourceAsStream("demoSheetData.json");
            byte[] bytes = new byte[resourceAsStream.available()];
            resourceAsStream.read(bytes);
            String jsonString = new String(bytes);
            exportToLocalFile(jsonString,"example");

        }catch (Exception e){
            e.printStackTrace();
        }
    }

}
