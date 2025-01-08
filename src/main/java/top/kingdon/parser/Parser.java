package top.kingdon.parser;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson2.JSONObject;
import com.alibaba.fastjson2.JSONReader;
import top.kingdon.parser.sheet.LuckySheet;

import java.io.IOException;
import java.io.InputStream;

public class Parser {
    public static LuckySheet parser(String jsonString){
        if(jsonString == null){
            return null;
        }
        LuckySheet luckySheet = JSONObject.parseObject(jsonString, LuckySheet.class, JSONReader.Feature.SupportSmartMatch);
        return luckySheet;

    }

    public static void main(String[] args){
        // 读取resources目录下sheetDataTemplate.json文件
        try{
            InputStream resourceAsStream = Parser.class.getClassLoader().getResourceAsStream("sheetDataTemplate.json");
            byte[] bytes = new byte[resourceAsStream.available()];
            resourceAsStream.read(bytes);
            String jsonString = new String(bytes);
            LuckySheet luckySheet = parser(jsonString);
            System.out.println(JSON.toJSONString(luckySheet));
        }catch (IOException e){
            e.printStackTrace();
        }

//        String jsonString = "{ \"name\": \"Cell\"}";
//        LuckySheetRowAndColumnLen luckySheetRowAndColumnLen =
//                JSONObject.parseObject(jsonString, LuckySheetRowAndColumnLen.class);

//        System.out.println(JSON.toJSONString(luckySheetRowAndColumnLen));


    }
}

