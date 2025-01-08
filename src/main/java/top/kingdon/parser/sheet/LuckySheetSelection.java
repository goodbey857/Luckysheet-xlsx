package top.kingdon.parser.sheet;

import lombok.Data;

@Data
public class LuckySheetSelection {
    int[] row; // selection start row and end row
    int[] column; // selection start column and end column
    int sheetIndex;
}
