package top.kingdon.parser.pivotTable;

import lombok.Data;
import top.kingdon.parser.sheet.LuckySheetSelection;

@Data
public class LuckySheetPivotTable {
    private LuckySheetSelection pivot_select_save; // Pivot table data source range
    private String pivotDataSheetIndex; // Data source sheet index, index is unique id
    private LuckySheetPivotTableField[] column; // Column area, include field
    private LuckySheetPivotTableField[] row; // Row area, include field
    private LuckySheetPivotTableField[] filter; // Filter area, include field
    private LuckySheetPivotTableFilterParam filterparm; // Save param after apply filter
    private LuckySheetPivotTableField[] values;
    private String showType;
    private Object[][] pivotDatas;
    private boolean drawPivotTable;
    private int[] pivotTableBoundary;
}
