package top.kingdon.parser.calcchain;

import lombok.Data;

@Data
public class LuckysheetCalcChain {
    private int r;
    private int c;
    private String index;
    private String func;
    private String color;
    private LuckysheetCalcChain parent;
    private LuckysheetCalcChain chidren;
    private int times;
}
