package top.kingdon.parser.dataVerification;

import lombok.Data;

import java.util.Map;

@Data
public class LuckysheetDataVerification {
    private Map<String, LuckysheetDataVerificationValue> dataVerification;

}
