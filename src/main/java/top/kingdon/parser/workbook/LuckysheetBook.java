package top.kingdon.parser.workbook;

import lombok.Data;
import top.kingdon.parser.sheet.LuckySheet;

import java.util.List;

@Data
public class LuckysheetBook {

    private String company;
    private String appVersion;
    private String creator;
    private String lastModifiedBy;
    private String createdTime;
    private String modifiedTime;
    private String fileName;
    private List<LuckySheet> sheets;
}
