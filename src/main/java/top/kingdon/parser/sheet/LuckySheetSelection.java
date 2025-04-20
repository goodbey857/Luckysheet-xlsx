package top.kingdon.parser.sheet;

import lombok.Data;

@Data
public class LuckySheetSelection {
    int left;
    int width;
    int top;
    int height;
    int left_move;
    int width_move;
    int top_move;
    int height_move;
    int[] row; // selection start row and end row
    int[] column; // selection start column and end column
    int sheetIndex;
    int row_focus;
    int column_focus;

}
