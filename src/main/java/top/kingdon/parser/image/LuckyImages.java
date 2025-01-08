package top.kingdon.parser.image;

import com.alibaba.fastjson2.annotation.JSONField;
import lombok.Data;

import java.util.Map;

@Data
public class LuckyImages {
    private String type; // 图片调整规则类型
    private String src; // 图片路径
    private int originWidth; // 图片原始宽度
    private int originHeight; // 图片原始高度
    // fastjson 反序列化时，如果字段名和json中的不一致
    @JSONField(name = "default")
    private DefaultPosition defaultPos; // 默认位置信息
    private CropConfig crop; // 图片裁剪信息
    private boolean isFixedPos; // 是否固定位置
    private int fixedLeft; // 固定左位移
    private int fixedTop; // 固定顶部位移
    private BorderConfig border; // 边框信息

    @Data
    public static class DefaultPosition {
        private int width;
        private int height;
        private int left;
        private int top;

    }
    @Data
    public static class CropConfig {
        private int width;
        private int height;
        private int offsetLeft;
        private int offsetTop;

    }
    @Data
    public static class BorderConfig {
        private int width;
        private String radius;
        private String style;
        private String color;

    }

}

