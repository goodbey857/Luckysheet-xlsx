package top.kingdon.parser.celldata;

import lombok.Data;

@Data
public class LuckyInlineString {
    private String ff; // font family
    private String fc; // font color
    private Integer fs; // font size
    private Integer cl; // strike
    private Integer un; // underline
    private Integer bl; // bold
    private Integer it; // italic
    private Integer va; // 1 for sub and 2 for super and 0 for none
    private String v;
}
