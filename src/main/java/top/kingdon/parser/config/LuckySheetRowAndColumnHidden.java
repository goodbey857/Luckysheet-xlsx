package top.kingdon.parser.config;

import lombok.Data;

import java.util.Map;

@Data
public class LuckySheetRowAndColumnHidden {
    private Map<String, Integer> hidden;
}
