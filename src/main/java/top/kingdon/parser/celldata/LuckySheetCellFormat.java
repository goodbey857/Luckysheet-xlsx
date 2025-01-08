package top.kingdon.parser.celldata;

import lombok.Data;

@Data
public class LuckySheetCellFormat {
    private String fa; // Format definition string
    private String t; // Cell Type
    private LuckyInlineString[] s;
}
