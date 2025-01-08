package top.kingdon.converter.cell;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.*;

import top.kingdon.parser.celldata.*;
import top.kingdon.utils.Util;

import java.util.Objects;
public class CellConverter {

    private final Workbook workbook;
    private final Sheet sheet;

    public CellConverter(Workbook workbook, Sheet sheet){
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void convert(LuckySheetCellData luckyCellData){
        if(Objects.isNull(luckyCellData)){ return;}

        LuckySheetCellDataValue luckyCell = luckyCellData.getV();
        if(Objects.isNull(luckyCell)){ return;}
        // 填数据
        Cell cell = fillValue(luckyCellData);
        // 填样式
        fillStyle(cell, luckyCell);
    }

    /**
     *  填充 字体、字体颜色 、加粗、斜体、字体大小、删除线、下划线 、垂直对齐、水平对齐、文本旋转、文本换行、引用前缀、背景色、文本格式
     * @param cell
     * @param luckyCell
     */
    private void fillStyle(Cell cell, LuckySheetCellDataValue luckyCell) {

        LuckySheetCellFormat cellFormat = luckyCell.getCt();
        String backgroundColor = luckyCell.getBg();

        Integer verticalAlignment = luckyCell.getVt();
        Integer horizontalAlignment = luckyCell.getHt();

        LuckySheetCellDataValueMerge cellMerge = luckyCell.getMc();

        Integer textRotation = luckyCell.getTr();
        Integer textWarp = luckyCell.getTb();
        Integer rotationText = luckyCell.getRt();
        Integer quotePrefix = luckyCell.getQp();

        CellStyle cellStyle = workbook.createCellStyle();

        // fontfamily, fontcolor, Bold, italic, font size, Cancelline, Underline
        Font font = fillCellFont(luckyCell);
        // Vertical alignment, 0 middle, 1 up, 2 down
        fillVerticalAlignment(verticalAlignment, cellStyle);
        // Horizontal alignment, 0 center, 1 left, 2 right
        fillHorizontalAlignment(horizontalAlignment, cellStyle);
        // Text rotation, 0: 0, 1: 45, 2: -45, 3 Vertical text, 4: 90, 5: -90
        fillTextRotation(textRotation, cellStyle);
        // Text wrap, 0 truncation, 1 overflow, 2 word wrap
        fillTextWrap(textWarp, cellStyle);
        // text rotation angle 0-180 alignment
        fillRotationText(cellStyle, rotationText);
        // quotePrefix, show number as string
        fillQuotePrefix(cellStyle, quotePrefix);
        // background,
        fillBackgroundColor(backgroundColor, cellStyle);

        // format
        XSSFRichTextString xssfRichTextString = fillFormat(cellFormat, cellStyle);
        if(Objects.nonNull(xssfRichTextString)){
            cell.setCellValue(xssfRichTextString);
        }

        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);

    }

    private XSSFRichTextString fillFormat(LuckySheetCellFormat cellFormat, CellStyle cellStyle) {
        if(Objects.isNull(cellFormat)){ return null;}

        String type = cellFormat.getT();

        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(cellFormat.getFa()));
        LuckyInlineString[] inlineStrings = cellFormat.getS();
//        处理inlineStyle
        if(type.equalsIgnoreCase("inlineStr") && Objects.nonNull(inlineStrings)){
            XSSFRichTextString xssfRichTextString = new XSSFRichTextString();
            for (LuckyInlineString inlineString : inlineStrings) {
                Font font = fillInlineStyleFont(inlineString);
                xssfRichTextString.append(inlineString.getV(), (XSSFFont) font);
            }
            return xssfRichTextString;
        }
        return null;
    }

    private void fillQuotePrefix(CellStyle cellStyle, Integer quotePrefix) {
        if(Objects.isNull(quotePrefix)){
            return;
        }
        cellStyle.setQuotePrefixed(quotePrefix == 1);


    }

    private void fillRotationText(CellStyle cellStyle, Integer rotationText) {
        if(Objects.isNull(rotationText)){
            return;
        }
        cellStyle.setRotation(rotationText.shortValue());


    }

    private void fillBackgroundColor(String backgroundColor, CellStyle cellStyle) {
        if(Objects.nonNull(backgroundColor)){
            byte[] rgbBytes = Util.toRGBBytes(backgroundColor.trim());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            ((XSSFCellStyle) cellStyle).setFillForegroundColor(new XSSFColor(rgbBytes));
        }
    }

    private void fillTextWrap(Integer textWarp, CellStyle cellStyle) {
        if(Objects.isNull(textWarp)){
            return;
        }
        switch (textWarp){
            case 0:
            case 1: cellStyle.setWrapText(false);break;
            case 2: cellStyle.setWrapText(true);break;
        }
    }

    private void fillTextRotation(Integer textRotation, CellStyle cellStyle) {
        if(Objects.isNull(textRotation)){
            return;
        }
        switch (textRotation){
            case 0: cellStyle.setRotation((short)0);break;
            case 1: cellStyle.setRotation((short)45);break;
            case 2: cellStyle.setRotation((short)-45);break;
            case 3: cellStyle.setRotation((short) 255);break;
            case 4: cellStyle.setRotation((short)90);break;
            case 5: cellStyle.setRotation((short)-90);break;
        }
    }

    private void fillHorizontalAlignment(Integer horizontalAlignment, CellStyle cellStyle) {
        if(Objects.isNull(horizontalAlignment)){
            return;
        }
        switch (horizontalAlignment){
            case 0: cellStyle.setAlignment(HorizontalAlignment.CENTER);break;
            case 1: cellStyle.setAlignment(HorizontalAlignment.LEFT);break;
            case 2: cellStyle.setAlignment(HorizontalAlignment.RIGHT);break;
        }
    }

    private void fillVerticalAlignment(Integer verticalAlignment, CellStyle cellStyle) {
        if(Objects.isNull(verticalAlignment)){
            return;
        }
        switch (verticalAlignment){
            case 0: cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);break;
            case 1: cellStyle.setVerticalAlignment(VerticalAlignment.TOP);break;
            case 2: cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);break;
        }
    }

    private Font fillInlineStyleFont(LuckyInlineString inlineStyle) {
        String fontFamily = inlineStyle.getFf();
        String fontColor = inlineStyle.getFc();
        Integer bold = inlineStyle.getBl();
        Integer italic = inlineStyle.getIt();
        Integer fontSize = inlineStyle.getFs();
        Integer cancelLine = inlineStyle.getCl();
        Integer underline = inlineStyle.getUn();
        return fillFont(fontFamily, fontColor, bold, italic, fontSize, cancelLine, underline);
    }

    private Font fillCellFont(LuckySheetCellDataValue luckyCell) {
        String fontFamily = luckyCell.getFf();
        String fontColor = luckyCell.getFc();
        Integer bold = luckyCell.getBl();
        Integer italic = luckyCell.getIt();
        Integer fontSize = luckyCell.getFs();
        Integer cancelline = luckyCell.getCl();
        Integer underline = luckyCell.getUn();

        return fillFont(fontFamily, fontColor, bold, italic, fontSize, cancelline, underline);

    }

    private Font fillFont(String fontFamily, String fontColor, Integer bold, Integer italic, Integer fontSize, Integer cancelline, Integer underline) {
        Font font = workbook.createFont();
        // fontfamily
        if(StringUtil.isNotBlank(fontFamily)){
            font.setFontName(convertFontFamily(fontFamily));
        }

        // color
        if(StringUtil.isNotBlank(fontColor)){
            byte[] rgbBytes = Util.toRGBBytes(fontColor.trim());
            ((XSSFFont) font).setColor(new XSSFColor(rgbBytes));
        }

        // bold
        if(Objects.nonNull(bold)){
            font.setBold(bold == 1);
        }

        // italic
        if(Objects.nonNull(italic)){
            font.setItalic(italic == 1);
        }
        // font size
        if(Objects.nonNull(fontSize)){
            font.setFontHeightInPoints(fontSize.shortValue());
        }

        // cancelline
        if(Objects.nonNull(cancelline)){
            font.setStrikeout(cancelline == 1);
        }
        // underline
        if (Objects.nonNull(underline)){
            switch(underline){
                case 1: font.setUnderline(Font.U_SINGLE);break;
                case 2: font.setUnderline(Font.U_DOUBLE);break;
                case 3: font.setUnderline(Font.U_SINGLE_ACCOUNTING);break;
                case 4: font.setUnderline(Font.U_DOUBLE_ACCOUNTING);break;
                case 0:
                default: font.setUnderline(Font.U_NONE);break;
            }
        }


        return font;
    }

//    0 Times New Roman、 1 Arial、2 Tahoma 、3 Verdana、4 微软雅黑、5 宋体（Song）、6 黑体（ST Heiti）、7 楷体（ST Kaiti）、
//    8 仿宋（ST FangSong）、9 新宋体（ST Song）、10 华文新魏、11 华文行楷、12 华文隶书
    private String convertFontFamily(String fontCode) {
        switch (fontCode){
            case "0": return "Times New Roman";
            case "1": return "Arial";
            case "2": return "Tahoma";
            case "3": return "Verdana";
            case "4": return "微软雅黑";
            case "5": return "宋体";
            case "6": return "黑体";
            case "7": return "楷体";
            case "8": return "仿宋";
            case "9": return "新宋体";
            case "10": return "华文新魏";
            case "11": return "华文行楷";
            case "12": return "华文隶书";
            default: return fontCode;
        }
    }


    private Cell fillValue(LuckySheetCellData luckyCellData) {
        if(Objects.isNull(luckyCellData)){
            return null;
        }

        LuckySheetCellDataValue luckyCell = luckyCellData.getV();
        String cellType = null;
        if(Objects.nonNull(luckyCell.getCt())){
            cellType = luckyCell.getCt().getT();
        }
        String value = luckyCell.getV();
        String formula = luckyCell.getF();

        Row row = sheet.getRow(luckyCellData.getR());
        if(Objects.isNull(row)){
            row = sheet.createRow(luckyCellData.getR());
        }
        Cell cell = row.createCell(luckyCellData.getC());

        if(StringUtil.isNotBlank(value)){
            switch(cellType){
                case "b": cell.setCellValue(Boolean.parseBoolean(value)); break;
                case "d": cell.setCellValue(Util.parseDateFrom1900(Double.parseDouble(value),false));break;
                case "n": cell.setCellValue(Double.parseDouble(value));break;
                case "inlineStr": break;
                case "str":
                case "s":
                case "e":
                default: cell.setCellValue(value);
            }
            cell.setCellValue(value);
        }else if(StringUtil.isNotBlank(formula)){
            // poi设置公式前面不要加等号‘=’
            cell.setCellFormula(formula.substring(1));
        }

        return cell;
    }
}
