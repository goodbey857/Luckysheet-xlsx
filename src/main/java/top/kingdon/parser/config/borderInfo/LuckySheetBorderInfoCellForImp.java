package top.kingdon.parser.config.borderInfo;

import lombok.Data;
import top.kingdon.parser.celldata.LuckySheetCellRange;
import top.kingdon.parser.sheet.LuckySheet;

import java.util.List;

@Data
public class LuckySheetBorderInfoCellForImp {
    private String rangeType;
    private List<String> cells;
    private LuckySheetBorderInfoCellValue value;
    private String borderType;
    private String style;
    private String color;
    private List<LuckySheetCellRange> range;


}
