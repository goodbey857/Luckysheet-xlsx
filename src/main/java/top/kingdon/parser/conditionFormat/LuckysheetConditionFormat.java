package top.kingdon.parser.conditionFormat;

import lombok.Data;
import top.kingdon.parser.sheet.LuckySheetSelection;

@Data
public class LuckysheetConditionFormat {
    private String type; // Option: default, databar, colorGradation, icons
    private LuckySheetSelection[] cellrange; // Valid range
    private Object format; // Style: String[] | IluckysheetCFDefaultFormat | IluckysheetCFIconsFormat
    private String conditionName; // Detailed settings, comparison parameters
    private LuckySheetSelection[] conditionRange; // Detailed settings, comparison range
    private Object[] conditionValue;

    @Data
    public static class LuckysheetCFDefaultFormat {
        private String textColor;
        private String cellColor;
    }

    @Data
    public static class LuckysheetCFIconsFormat {
        private Integer len;
        private Integer leftMin;
        private Integer top;

    }
}
