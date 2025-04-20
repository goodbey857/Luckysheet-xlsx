package top.kingdon.parser.sheet;

import com.alibaba.fastjson.annotation.JSONField;
import lombok.Data;
import top.kingdon.parser.calcchain.LuckysheetCalcChain;
import top.kingdon.parser.celldata.LuckySheetCellData;
import top.kingdon.parser.conditionFormat.LuckysheetConditionFormat;
import top.kingdon.parser.config.LuckySheetConfig;
import top.kingdon.parser.dataVerification.LuckysheetDataVerification;
import top.kingdon.parser.hyperlink.LuckysheetHyperlink;
import top.kingdon.parser.image.LuckyImages;
import top.kingdon.parser.pivotTable.LuckySheetPivotTable;

import java.util.List;
import java.util.Map;

@Data
public class LuckySheet {
    private String name;
    private String color;
    private LuckySheetConfig config;
    private String index;
    private String status;
    private String order;
    private int row;
    private int column;
    private LuckySheetSelection[] luckysheet_select_save;
    private double scrollLeft;
    private double scrollTop;
    private double zoomRatio;
    private String showGridLines;
    private Integer defaultColWidth;
    private Integer defaultRowHeight;
    private LuckySheetCellData[] celldata;
    private LuckySheetChart[] chart;
    private LuckySheetPivotTable pivotTable;
    private boolean isPivotTable;
    private LuckysheetConditionFormat[] luckysheet_conditionformat_save;
    private LuckysheetFrozen frozen;
    private LuckysheetCalcChain[] calcChain;
    private Map<String,LuckyImages> images;
    private LuckysheetDataVerification dataVerification;
    private LuckysheetHyperlink hyperlink;
    private int hide;
}
