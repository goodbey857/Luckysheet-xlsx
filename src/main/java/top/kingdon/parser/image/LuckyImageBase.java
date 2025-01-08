package top.kingdon.parser.image;

import lombok.Data;

@Data
public class LuckyImageBase {
    private LuckyImageBorder border;
    private LuckyImageCrop crop;
    private LuckyImageDefault _default;

    private double fixedLeft;
    private double fixedTop;
    private boolean isFixedPos;
    private double originHeight;
    private double originWidth;
    private String src;
    private String type;
}
