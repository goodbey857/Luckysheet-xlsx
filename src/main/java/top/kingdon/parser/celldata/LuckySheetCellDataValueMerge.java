package top.kingdon.parser.celldata;

import lombok.Data;

@Data
public class LuckySheetCellDataValueMerge {
    private Integer rs; //row of merge cell length, only main merge cell, every merge cell has only one main mergeCell
    private Integer cs; //column of merge cell length, only main merge cell, every merge cell has only one main mergeCell
    private int r; //main merge cell row Number, other cell link to main cell
    private int c; //main merge cell column Number, other cell link to main cell

}
