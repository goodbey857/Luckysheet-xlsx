package top.kingdon.parser.dataVerification;

import lombok.Data;

@Data
public class LuckysheetDataVerificationValue {
    private LuckysheetDataVerificationType type;
    private String type2;
    private Object value1;
    private Object value2;
    private boolean checked;
    private boolean remote;
    private boolean prohibitInput;
    private boolean hintShow;
    private String hintText;
}
