package top.kingdon.parser.celldata;

import lombok.Data;

@Data
public class LuckySheetCellData {
    private int r; // cell row number
    private int c; // cell column number
    private LuckySheetCellDataValue v; // cell value (IluckySheetCelldataValue, String, or null)

}
