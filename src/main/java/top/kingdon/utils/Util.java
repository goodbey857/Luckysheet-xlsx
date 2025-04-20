package top.kingdon.utils;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Util {


    public static  byte[] toRGBBytes(String fontColor) {
        Pattern pattern = Pattern.compile("rgb\\((\\d+),\\s*(\\d+),\\s*(\\d+)\\)");
        Matcher matcher = pattern.matcher(fontColor);

        // #0000ff
        Pattern pattern2 = Pattern.compile("#([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})");
        Matcher matcher2 = pattern2.matcher(fontColor);
        byte[] bytes = new byte[3];
        if(matcher.find()){
            bytes[0] = (byte) Integer.parseInt(matcher.group(1));
            bytes[1] = (byte) Integer.parseInt(matcher.group(2));
            bytes[2] = (byte) Integer.parseInt(matcher.group(3));

        } else if (matcher2.find()){
            bytes[0] = (byte) Integer.parseInt(matcher2.group(1), 16);
            bytes[1] = (byte) Integer.parseInt(matcher2.group(2), 16);
            bytes[2] = (byte) Integer.parseInt(matcher2.group(3), 16);
        }
        return bytes;
    }


    public static byte[] decodeBase64(String input) {
        byte[] bytes = null;
        String base64Data = getPureBase64String(input);
        if (base64Data == null) return null;

        // 5. 尝试解码验证
        try {
            bytes = Base64.getDecoder().decode(base64Data);
            return bytes;
        } catch (IllegalArgumentException e) {
            return null;
        }
    }

    public static boolean isValidBase64String(String input) {
        return getPureBase64String(input) != null;
    }

    private static String getPureBase64String(String input) {
        // 1. 快速排除空字符串
        if (input == null || input.isEmpty()) {
            return null;
        }

        // 2. 如果是 Base64 数据 URI，移除前缀部分
        String base64Data = input;
        if (input.startsWith("data:")) {
            int commaIndex = input.indexOf(",");
            if (commaIndex == -1 || commaIndex + 1 >= input.length()) {
                return null;
            }
            base64Data = input.substring(commaIndex + 1); // 提取实际 Base64 数据
        }

        // 3. 长度检查：必须是4的倍数
        if (base64Data.length() % 4 != 0) {
            return null;
        }

        // 4. 正则匹配合法的 Base64 字符集
        if (!base64Data.matches("^[A-Za-z0-9+/]*={0,2}$")) {
            return null;
        }
        return base64Data;
    }

    public static List<String> splitJsonArray(String jsonArray) {
        if (jsonArray == null || jsonArray.isEmpty()) {
            return null;
        }
        ArrayList<String> result = new ArrayList<>();
        Deque<Integer> deque = new ArrayDeque<>();
        for( int i = 0; i < jsonArray.length(); i++) {
            char c = jsonArray.charAt(i);
            if (c == '{') {
                deque.push(i);
            } else if (c == '}') {
                Integer startIndex = deque.pop();
                if (deque.isEmpty()) {
                    String json = jsonArray.substring(startIndex, i + 1);
                    result.add(json);
                }
            }

        }
        return result;
    }

}
