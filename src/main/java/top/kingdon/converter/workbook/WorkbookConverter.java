package top.kingdon.converter.workbook;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.kingdon.converter.sheet.SheetConverter;
import top.kingdon.parser.sheet.LuckySheet;
import top.kingdon.parser.workbook.LuckysheetBook;

public class WorkbookConverter {
    public WorkbookConverter(LuckysheetBook luckysheetBook){
        XSSFWorkbook workbook = new XSSFWorkbook();
        for(LuckySheet sheet : luckysheetBook.getSheets()){
            SheetConverter sheetConverter = new SheetConverter(workbook);
            sheetConverter.convert(sheet);
        }
    }
}
