package top.kingdon.parser.hyperlink;

import lombok.Data;

@Data
public class LuckysheetHyperlink {
    private String linkAddress;
    private String linkTooltip;
    private LuckysheetHyperlinkType linkType;
    private String display;
}
