package top.kingdon.parser.celldata;

import lombok.Data;

@Data
public class LuckySheetCellDataValue {
    private LuckySheetCellFormat ct; // celltype, Cell value format: text, time, etc.
    private String bg; // background, #fff000
    private String ff; // fontfamily
    private String fc; // fontcolor
    private Integer bl; // Bold
    private Integer it; // italic
    private Integer fs; // font size
    private Integer cl; // Cancelline, 0 Regular, 1 Cancelline
    private Integer un; // underline, 0 Regular, 1 underlines, fonts
    private Integer vt; // Vertical alignment, 0 middle, 1 up, 2 down
    private Integer ht; // Horizontal alignment, 0 center, 1 left, 2 right
    private LuckySheetCellDataValueMerge mc; // Merge Cells
    private Integer tr; // Text rotation, 0: 0, 1: 45, 2: -45, 3 Vertical text, 4: 90, 5: -90
    private Integer tb; // Text wrap, 0 truncation, 1 overflow, 2 word wrap
    private String v; // Original value
    private String m; // Display value
    private Integer rt; // text rotation angle 0-180 alignment
    private String f; // formula
    private Integer qp; // quotePrefix, show number as string
}
