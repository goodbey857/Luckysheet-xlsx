package top.kingdon.parser.celldata;

import lombok.Data;

@Data
public class LuckySheetCellRange {
    private int[] row;
    private int[] column;

    public int getRowStart() {
        return row[0];
    }
    public int getRowEnd() {
        return row[1];
    }
    public int getColumnStart() {
        return column[0];
    }
    public int getColumnEnd() {
        return column[1];
    }
}
