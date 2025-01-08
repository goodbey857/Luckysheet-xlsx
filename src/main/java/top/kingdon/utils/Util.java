package top.kingdon.utils;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Base64;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Util {


    public static Date parseDateFrom1900(double serial, boolean date1904) {
        // 定义基准时间（1899年12月31日 00:00:00 UTC）
        Calendar dnthreshUtc = Calendar.getInstance();
        dnthreshUtc.set(1899, Calendar.DECEMBER, 31, 0, 0, 0);
        dnthreshUtc.set(Calendar.MILLISECOND, 0);
        long dnthreshUtcMillis = dnthreshUtc.getTimeInMillis();

        // 计算序列号对应的毫秒数
        long epoch = (long) (serial * 24 * 60 * 60 * 1000) + dnthreshUtcMillis;

        // 如果使用1904年日期系统，调整时间
        if (date1904) {
            epoch += 1461L * 24 * 60 * 60 * 1000;
        } else {
            // 定义1900年3月1日
            Calendar base1904 = Calendar.getInstance();
            base1904.set(1900, Calendar.MARCH, 1, 0, 0, 0);
            base1904.set(Calendar.MILLISECOND, 0);

            // 如果日期大于或等于1900年3月1日，调整时间
            if (new Date(epoch).after(base1904.getTime())) {
                epoch -= 24L * 60 * 60 * 1000;
            }
        }

        // 返回对应的日期对象
        return new Date(epoch);
    }

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

}
