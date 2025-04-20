package top.kingdon.parser.sheet;

import lombok.Data;

@Data
public class LuckysheetFrozen {
    private String type;
    private Range range;

    @Data
    public static class Range{
        private Integer row_focus;
        private Integer column_focus;
    }

}


