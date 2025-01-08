package top.kingdon.parser.config;

import lombok.Data;
import top.kingdon.parser.config.borderInfo.LuckySheetBorderInfoCellForImp;
import top.kingdon.parser.config.merge.LuckySheetConfigMerge;
import top.kingdon.parser.config.merge.LuckySheetConfigMerges;

import java.util.List;
import java.util.Map;

@Data
public class LuckySheetConfig {
    private Map<String, LuckySheetConfigMerge> merge; //merge handler
    private List<LuckySheetBorderInfoCellForImp> borderInfo; //range border
    private Map<String, Integer> rowlen; //every row's height
    private Map<String, Integer> columnlen; //every column's width
    private Map<String, Integer> rowhidden; //setting be hidden rows
    private Map<String, Integer> colhidden; //setting be hidden columns

    private Map<String, Integer> customHeight; //user operate row height
    private Map<String, Integer> customWidth; //user operate column width

}
