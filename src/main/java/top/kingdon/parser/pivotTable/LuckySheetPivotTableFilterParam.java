package top.kingdon.parser.pivotTable;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

@Data
public class LuckySheetPivotTableFilterParam {
    Map<String, LuckySheetPivotTableFilterParamItem> map = new HashMap<>();

}
