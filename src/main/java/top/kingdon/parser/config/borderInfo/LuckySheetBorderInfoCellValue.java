package top.kingdon.parser.config.borderInfo;

import lombok.Data;

@Data
public class LuckySheetBorderInfoCellValue {
    private int rowIndex;
    private int colIndex;
    private LuckySheetBorderInfoCellValueStyle l;
    private LuckySheetBorderInfoCellValueStyle r;
    private LuckySheetBorderInfoCellValueStyle t;
    private LuckySheetBorderInfoCellValueStyle b;
}
