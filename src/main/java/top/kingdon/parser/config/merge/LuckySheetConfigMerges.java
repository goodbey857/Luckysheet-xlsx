package top.kingdon.parser.config.merge;

import lombok.Data;

import java.util.Map;

@Data
public class LuckySheetConfigMerges {
    private Map<String, LuckySheetConfigMerge> merges;
}
