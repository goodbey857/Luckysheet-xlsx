package top.kingdon.parser;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson2.JSONObject;
import com.alibaba.fastjson2.JSONReader;
import top.kingdon.parser.sheet.LuckySheet;

import java.io.IOException;
import java.io.InputStream;

/**
 * 负责将字符串解析成LuckySheet对象。
 */
public class Parser {
    public static LuckySheet parser(String jsonString) {
        if (jsonString == null) {
            return null;
        }
        LuckySheet luckySheet = null;
        try {
            luckySheet = JSONObject.parseObject(jsonString, LuckySheet.class, JSONReader.Feature.SupportSmartMatch);

        }catch (Exception e){
            e.printStackTrace();
        }
        return luckySheet;

    }

}

