package top.kingdon.converter.sheet;

import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import top.kingdon.converter.cell.CellConverter;
import top.kingdon.parser.celldata.LuckySheetCellData;
import top.kingdon.parser.celldata.LuckySheetCellRange;
import top.kingdon.parser.conditionFormat.LuckysheetConditionFormat;
import top.kingdon.parser.config.LuckySheetConfig;
import top.kingdon.parser.config.borderInfo.LuckySheetBorderInfoCellForImp;
import top.kingdon.parser.config.borderInfo.LuckySheetBorderInfoCellValue;
import top.kingdon.parser.config.borderInfo.LuckySheetBorderInfoCellValueStyle;
import top.kingdon.parser.config.merge.LuckySheetConfigMerge;
import top.kingdon.parser.image.LuckyImages;
import top.kingdon.parser.sheet.LuckySheet;
import top.kingdon.parser.sheet.LuckySheetSelection;
import top.kingdon.parser.sheet.LuckysheetFrozen;
import top.kingdon.utils.Util;

import java.util.*;

import static top.kingdon.utils.Util.decodeBase64;

public class SheetConverter {
    private Workbook workbook;

    public SheetConverter(){
        this.workbook = new XSSFWorkbook();
    };
    public SheetConverter(Workbook workbook){
        this.workbook = workbook;
    }


    /**
     * 设置单元格合并、设置单元格内容、设置边框、设置行宽，列高、设置隐藏行，隐藏列
     * @param luckySheet
     */
    public void convert(LuckySheet luckySheet){

        LuckySheetConfig config = luckySheet.getConfig();
        Map<String, LuckySheetConfigMerge> merge = config.getMerge();
        Map<String, Integer> columnlen = config.getColumnlen();
        Map<String, Integer> rowlen = config.getRowlen();
        Map<String, Integer> colhidden = config.getColhidden();
        Map<String, Integer> rowhidden = config.getRowhidden();
        LuckySheetCellData[] celldata = luckySheet.getCelldata();

        Sheet sheet = workbook.createSheet(luckySheet.getName());
//        PIXEL to Point
        if(Objects.nonNull(luckySheet.getDefaultRowHeight())){
            short rowHeight = calcRowHeight(luckySheet.getDefaultRowHeight());
            sheet.setDefaultRowHeight(rowHeight);
        }
        if(Objects.nonNull(luckySheet.getDefaultColWidth())){
            int colWidth = calcColWidth(luckySheet.getDefaultColWidth());
            sheet.setDefaultColumnWidth(colWidth);
        }


        // 设置单元格合并 merge
        mergeCell(sheet, merge);

        // 设置单元格内容
        fillCellData(sheet, celldata);

        // 设置边框
        setBorder(sheet, luckySheet.getConfig().getBorderInfo());

        // 设置行宽，列高
        setRowHeight(sheet, rowlen);
        setColumnWidth(sheet, columnlen);

        // 设置隐藏行，隐藏列
        setRowHidden(sheet, rowhidden);
        setColumnHidden(sheet, colhidden);

        // 冻结行、冻结列
        setFreeze(sheet, luckySheet.getFrozen());

        // 填充条件格式
        fillConditionalFormatting(sheet, luckySheet.getLuckysheet_conditionformat_save());

        // 填充图片
        fillImages(sheet, luckySheet.getImages(),luckySheet.getConfig(),luckySheet.getDefaultColWidth(),luckySheet.getDefaultRowHeight());

    }

    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * default
     * "dataBar": 数据条
     * "icons": 图标集
     * "colorGradation": 色阶
     * @param sheet
     * @param luckysheetConditionformatSave
     */
    private void fillConditionalFormatting(Sheet sheet, LuckysheetConditionFormat[] luckysheetConditionformatSave) {
        if(Objects.isNull(luckysheetConditionformatSave) || luckysheetConditionformatSave.length == 0){
            return;
        }
        for (LuckysheetConditionFormat luckysheetConditionformat : luckysheetConditionformatSave) {
            switch (luckysheetConditionformat.getType()){
                case "dataBar": fillDataBarConditionalFormatting(sheet, luckysheetConditionformat);break;
                case "icons": fillIconsConditionalFormatting(sheet,luckysheetConditionformat);break;
                case "colorGradation": fillColorGradationConditionalFormatting(sheet,luckysheetConditionformat); break;
                case "default": fillDefaultConditionalFormatting(sheet,luckysheetConditionformat); break;
                default:break;
            }
        }

    }

    private void fillIconsConditionalFormatting(Sheet sheet, LuckysheetConditionFormat luckysheetConditionformat) {
        LuckySheetSelection[] cellRange = luckysheetConditionformat.getCellrange();

        LuckysheetConditionFormat.LuckysheetCFIconsFormat iconsFormat =
                ((JSONObject) luckysheetConditionformat.getFormat()).
                        toJavaObject(LuckysheetConditionFormat.LuckysheetCFIconsFormat.class);
        CellRangeAddress[] cellRangeAddresses = getCellRangeAddresses(cellRange);
        SheetConditionalFormatting sheetSF = sheet.getSheetConditionalFormatting();
        IconMultiStateFormatting.IconSet iconSet = formatIconSet(iconsFormat);
        if(Objects.isNull(iconSet)){
            return;
        }
        ConditionalFormattingRule rule = sheetSF.createConditionalFormattingRule(iconSet);

        sheetSF.addConditionalFormatting(cellRangeAddresses,rule);



    }

    private IconMultiStateFormatting.IconSet formatIconSet(LuckysheetConditionFormat.LuckysheetCFIconsFormat iconsFormat){
        String code = String.format("%d_%d_%d",iconsFormat.getLen(),iconsFormat.getLeftMin(),iconsFormat.getTop());
        switch(code){
            case "3_0_0": return IconMultiStateFormatting.IconSet.GYR_3_ARROW;
            case "3_5_0": return IconMultiStateFormatting.IconSet.GREY_3_ARROWS;
            case "3_0_1": return IconMultiStateFormatting.IconSet.GYR_3_ARROW;
            case "4_5_1": return IconMultiStateFormatting.IconSet.GREY_4_ARROWS;
            case "4_0_2": return IconMultiStateFormatting.IconSet.GYR_4_ARROWS;
            case "5_5_2": return IconMultiStateFormatting.IconSet.GREY_5_ARROWS;
            case "5_0_3": return IconMultiStateFormatting.IconSet.GYYYR_5_ARROWS;
            case "3_0_4": return IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS;
            case "3_5_4": return IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS_BOX;
            case "3_0_5": return IconMultiStateFormatting.IconSet.GYR_3_SHAPES;
            case "4_5_5": return IconMultiStateFormatting.IconSet.GYRB_4_TRAFFIC_LIGHTS;
            case "4_0_6": return IconMultiStateFormatting.IconSet.RB_4_TRAFFIC_LIGHTS;
            case "3_0_7": return IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS_CIRCLE;
            case "3_5_7": return IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS;
            case "3_0_8": return IconMultiStateFormatting.IconSet.GYR_3_FLAGS;
            case "3_0_9": return null;
            case "4_5_9": return IconMultiStateFormatting.IconSet.RATINGS_4;
            case "5_0_10": return IconMultiStateFormatting.IconSet.QUARTERS_5;
            case "5_5_10": return IconMultiStateFormatting.IconSet.RATINGS_5;
            case "5_0_11": return IconMultiStateFormatting.IconSet.QUARTERS_5;
            default: return null;
        }
    }

    private void fillDataBarConditionalFormatting(Sheet sheet, LuckysheetConditionFormat luckysheetConditionformat) {
        LuckySheetSelection[] cellRange = luckysheetConditionformat.getCellrange();
        String[] format = ((JSONArray) luckysheetConditionformat.getFormat()).toArray(String.class);
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule = sheetCF.createConditionalFormattingRule(
                new XSSFColor(Util.toRGBBytes(format[0])));
        XSSFDataBarFormatting dataBar = (XSSFDataBarFormatting) rule.getDataBarFormatting();

        dataBar.setColor(new XSSFColor(Util.toRGBBytes("#00ff00")));

        CellRangeAddress[] regions = getCellRangeAddresses(cellRange);

        sheetCF.addConditionalFormatting(regions, rule);
    }

    private static CellRangeAddress[] getCellRangeAddresses(LuckySheetSelection[] cellRange) {
        CellRangeAddress[] regions = new CellRangeAddress[cellRange.length];
        for (int i = 0; i < cellRange.length; i++) {
            int startRow = cellRange[i].getRow()[0];
            int endRow = cellRange[i].getRow()[1];
            int startCol = cellRange[i].getColumn()[0];
            int endCol = cellRange[i].getColumn()[1];
            regions[i] = new CellRangeAddress(startRow, endRow, startCol, endCol);
        }
        return regions;
    }


    private void fillColorGradationConditionalFormatting(Sheet sheet, LuckysheetConditionFormat luckysheetConditionformat) {
       LuckySheetSelection[] cellRange = luckysheetConditionformat.getCellrange();
       String[] format = ((JSONArray) luckysheetConditionformat.getFormat()).toArray(String.class);
       CellRangeAddress[] regions = Arrays.stream(cellRange)
               .map(range -> new CellRangeAddress(range.getRow()[0], range.getRow()[1], range.getColumn()[0], range.getColumn()[1]))
               .toArray(CellRangeAddress[]::new);


       SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
       ConditionalFormattingRule rule = sheetCF.createConditionalFormattingColorScaleRule();
       // 获取颜色标度格式化对象
       ColorScaleFormatting colorScale = rule.getColorScaleFormatting();
       // 设置色阶的颜色
       XSSFColor[] colors = new XSSFColor[format.length];
       for (int i = 0; i < format.length; i++) {
           colors[i] = new XSSFColor(Util.toRGBBytes(format[i]));
       }
       colorScale.setColors(colors);
       // 设置三色阶的阈值
       ConditionalFormattingThreshold[] thresholds = new ConditionalFormattingThreshold[format.length];
       if(format.length == 2){
           // 最小值
           thresholds[0] = colorScale.createThreshold();
           thresholds[0].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
           thresholds[0].setValue(0.0);
           // 最大值
           thresholds[1] = colorScale.createThreshold();
           thresholds[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
           thresholds[1].setValue(100.0);
       } else if (format.length == 3) {
           // 最小值
           thresholds[0] = colorScale.createThreshold();
           thresholds[0].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
           thresholds[0].setValue(0.0);
           // 中间值
           thresholds[1] = colorScale.createThreshold();
           thresholds[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
           thresholds[1].setValue(50.0);
           // 最大值
           thresholds[2] = colorScale.createThreshold();
           thresholds[2].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENT);
           thresholds[2].setValue(100.0);
       }

       colorScale.setThresholds(thresholds);
       // 应用条件格式
       sheetCF.addConditionalFormatting(regions, rule);


   }



    /**
     * conditionName的取值如下
     * greaterThan
     * lessThan
     * betweenness
     * equal
     * textContains
     * occurrenceDate
     * duplicateValue
     *
     * top10%
     * top10
     * bottom10%
     * bottom10
     * SubAverage
     * AboveAverage
     * formula
     * @param sheet
     * @param luckysheetConditionformat
     */
    private void fillDefaultConditionalFormatting(Sheet sheet, LuckysheetConditionFormat luckysheetConditionformat) {



        // 获取条件格式管理器
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // 应用范围
        LuckySheetSelection[] cellRanges = luckysheetConditionformat.getCellrange();
        CellRangeAddress[] regions = Arrays.stream(cellRanges)
                .map(range -> new CellRangeAddress(range.getRow()[0], range.getRow()[1], range.getColumn()[0], range.getColumn()[1]))
                .toArray(CellRangeAddress[]::new);

        // 获取第一个范围的起始和结束列
        LuckySheetSelection firstRange = cellRanges[0];
        String startColumn = CellReference.convertNumToColString(firstRange.getColumn()[0]);
        String endColumn = CellReference.convertNumToColString(firstRange.getColumn()[1]);
        int startRow = firstRange.getRow()[0] + 1;
        int endRow = firstRange.getRow()[1] + 1;
        String rangeFormula = String.format("%s%d:%s%d", startColumn, startRow, endColumn, endRow);
        Object[] conditionValue = luckysheetConditionformat.getConditionValue();

        // 创建条件格式规则
        String conditionName = luckysheetConditionformat.getConditionName();
        ConditionalFormattingRule rule = null;

        switch (conditionName) {
            case "greaterThan":
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.GT,
                        String.valueOf(conditionValue[0]));
                break;
            case "lessThan":
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.LT,
                        String.valueOf(conditionValue[0]));
                break;
            case "betweenness":
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.BETWEEN,
                        String.valueOf(conditionValue[0]),
                        String.valueOf(conditionValue[1])
                );
                break;
            case "equal":
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.EQUAL,
                        String.valueOf(conditionValue[0]));
                break;
            case "textContains":
                rule = sheetCF.createConditionalFormattingRule(
                        String.format("SEARCH(\"%s\",%s)>0", conditionValue[0], rangeFormula));
                break;
            case "occurrenceDate":
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.BETWEEN, String.valueOf(conditionValue[0]));
                break;
            case "duplicateValue":
                if("1".equals(conditionValue[0])){
                    rule = sheetCF.createConditionalFormattingRule(
                            String.format("COUNTIF(%s,%s%d)=1", rangeFormula, startColumn,startRow));
                }else{
                    rule = sheetCF.createConditionalFormattingRule(
                            String.format("COUNTIF(%s,%s%d)>1", rangeFormula, startColumn,startRow));
                }

                break;

            case "top10":
                rule = sheetCF.createConditionalFormattingRule(
                        String.format("RANK(%s1,%s)<=10", startColumn, rangeFormula));
                break;
            case "top10%":
                rule = sheetCF.createConditionalFormattingRule(
                        String.format("RANK(%s1,%s)/COUNTA(%s)<=0.1", startColumn, rangeFormula, rangeFormula));
                break;
            case "bottom10":
                rule = sheetCF.createConditionalFormattingRule(
                        String.format("RANK(%s1,%s)>=COUNTA(%s)-9", startColumn, rangeFormula, rangeFormula));
                break;
            case "bottom10%":
                rule = sheetCF.createConditionalFormattingRule(
                        String.format("RANK(%s1,%s)/COUNTA(%s)>=0.9", startColumn, rangeFormula, rangeFormula));
                break;
            case "SubAverage":
//                rule = sheetCF.createConditionalFormattingRule(String.format("%s1<AVERAGE(%s)", startColumn, rangeFormula));
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.LT,String.format("AVERAGE(%s)", rangeFormula));
                break;
            case "AboveAverage":
//                rule = sheetCF.createConditionalFormattingRule(String.format("%s1>AVERAGE(%s)", startColumn, rangeFormula));
                rule = sheetCF.createConditionalFormattingRule(
                        ComparisonOperator.GT,String.format("AVERAGE(%s)", rangeFormula));
                break;
            case "formula":
                rule = sheetCF.createConditionalFormattingRule(
                        conditionValue[0].toString());
                break;
            default:
                break;
        }

        // 设置格式
        if (rule != null) {
            FontFormatting fontFormatting = rule.createFontFormatting();
            JSONObject luckysheetCFDefaultFormat =
                    (JSONObject) luckysheetConditionformat.getFormat();

            if (luckysheetCFDefaultFormat.getString("textColor") != null) {
                byte[] rgbBytes = Util.toRGBBytes(luckysheetCFDefaultFormat.getString("textColor"));
                XSSFColor textColor = new XSSFColor(rgbBytes);
//                XSSFColor textColor =
//                        getColorFromHex(luckysheetConditionformat.getFormat().getTextColor());
                fontFormatting.setFontColor(textColor);
            }

            PatternFormatting patternFormatting = rule.createPatternFormatting();
            if (luckysheetCFDefaultFormat.getString("cellColor") != null) {
                byte[] rgbBytes = Util.toRGBBytes(luckysheetCFDefaultFormat.getString("cellColor"));
                XSSFColor cellColor =new XSSFColor(rgbBytes);
                patternFormatting.setFillBackgroundColor(cellColor);
            }

            // 添加条件格式规则到工作表
            sheetCF.addConditionalFormatting(regions, rule);
        }
    }


    /**
     * "row": 冻结首行
     * "column": 冻结首列
     * "both": 冻结行列
     * "rangeRow": 冻结行到选区
     * "rangeColumn": 冻结列到选区
     * "rangeBoth": 冻结行列到选区
     * "cancel": 取消冻结
     * @param sheet
     * @param frozen
     */
    private void setFreeze(Sheet sheet, LuckysheetFrozen frozen) {
        if(Objects.isNull(frozen)){
            return;
        }
        switch (frozen.getType()){
            case "row": sheet.createFreezePane(0,1);break;
            case "column": sheet.createFreezePane(1,0);break;
            case "both": sheet.createFreezePane(1,1);break;
            case "rangeRow": sheet.createFreezePane(0,frozen.getRange().getRow_focus()); break;
            case "rangeColumn": sheet.createFreezePane(frozen.getRange().getColumn_focus(),0); break;
            case "rangeBoth": sheet.createFreezePane(frozen.getRange().getColumn_focus(),frozen.getRange().getRow_focus()); break;
            case "cancel":
            default:break;

        }



    }

    private void fillImages(Sheet sheet, Map<String,LuckyImages> luckyImagesList, LuckySheetConfig config,Integer defaultColWidth, Integer defaultRowHeight) {
        if(Objects.isNull(luckyImagesList) || luckyImagesList.isEmpty()){
            return;
        }
        if(Objects.isNull(defaultColWidth)){
            defaultColWidth = 73;
        }
        if(Objects.isNull(defaultRowHeight)){
            defaultRowHeight = 20;
        }

        Map<String, Integer> rowlen = config.getRowlen();
        Map<String, Integer> columnlen = config.getColumnlen();

        for (   Map.Entry<String, LuckyImages> luckyImagesEntrySet : luckyImagesList.entrySet()) {
            LuckyImages luckyImages = luckyImagesEntrySet.getValue();
            if(Objects.isNull(luckyImages)){
                continue;
            }
            String src = luckyImages.getSrc();
            // 判断scr是不是base64编码
            int pictureIndex = 0;
            if (Util.isValidBase64String(src)) {
                byte[] bytes = decodeBase64(src);
                pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            }
            // 创建绘图容器
            Drawing<?> drawing =  sheet.createDrawingPatriarch();

            CreationHelper creationHelper = workbook.getCreationHelper();

            // 定义图片锚点
            ClientAnchor anchor = creationHelper.createClientAnchor();


            if (luckyImages.isFixedPos()) {
                Integer[] x = calcRowCol(luckyImages.getFixedLeft(), defaultColWidth, columnlen);
                Integer[] y = calcRowCol(luckyImages.getFixedTop(), defaultRowHeight, rowlen);
                // 固定位置
                anchor.setCol1(x[0]);
                anchor.setDx1(Units.pixelToEMU(x[1]));
                anchor.setRow1(y[0]);
                anchor.setDy1(Units.pixelToEMU(y[1]));


            } else if(Objects.nonNull(luckyImages.getDefaultPos())) {
                Integer[] x = calcRowCol(luckyImages.getDefaultPos().getLeft(), defaultColWidth, columnlen);
                Integer[] y = calcRowCol(luckyImages.getDefaultPos().getTop(), defaultRowHeight, rowlen);
                // 相对位置
                anchor.setCol1(x[0]);
                anchor.setDx1(Units.pixelToEMU(x[1]));
                anchor.setRow1(y[0]);
                anchor.setDy1(Units.pixelToEMU(y[1]));


            }
//             图片大小
            if(Objects.nonNull(luckyImages.getCrop())){
                int width = luckyImages.getCrop().getWidth();
                int height = luckyImages.getCrop().getHeight();
                Integer[] x = calcRowCol(width, defaultColWidth, columnlen);
                Integer[] y = calcRowCol(height, defaultRowHeight, rowlen);
                anchor.setCol2(y[0]);
                anchor.setRow2(x[0]);
                anchor.setDx2(Units.pixelToEMU(x[1]));
                anchor.setDy2(Units.pixelToEMU(y[1]));

            }

            // 配置移动与调整规则
            switch (luckyImages.getType()) {
                case "1":
                    anchor.setAnchorType(XSSFClientAnchor.AnchorType.MOVE_AND_RESIZE);
                    break;
                case "2":
                    anchor.setAnchorType(XSSFClientAnchor.AnchorType.MOVE_DONT_RESIZE);
                    break;
                case "3":
                default:
                    anchor.setAnchorType(XSSFClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                    break;
            }


            // 插入图片
            Picture picture = drawing.createPicture(anchor, pictureIndex);
            picture.resize();
        }

    }

    private Integer[]  calcRowCol(int len, int defaultLen, Map<String,Integer> lenMap){
        int currRowCol = 0;
        Integer rowColLen = lenMap.getOrDefault(String.valueOf(currRowCol),defaultLen);
        while(len > rowColLen){
            len-= rowColLen;
            rowColLen = lenMap.getOrDefault(String.valueOf(++currRowCol),defaultLen);
        }
        return new Integer[]{currRowCol,len};
    }




    /**
     * @param sheet
     * @param borderInfoList
     */
    private void setBorder(Sheet sheet, List<LuckySheetBorderInfoCellForImp> borderInfoList) {
        if(Objects.isNull(borderInfoList)){
            return;
        }
        for (LuckySheetBorderInfoCellForImp borderInfo: borderInfoList) {
            if(borderInfo.getRangeType().equals("cell")){
                cellBorder(sheet, borderInfo);
            } else if (borderInfo.getRangeType().equals("range")) {
                rangeBorder(sheet, borderInfo);
            }
        }
    }

    /**
     * borderType："border-left" | "border-right" | "border-top" | "border-bottom" | "border-all"
     * | "border-outside" | "border-inside" | "border-horizontal" | "border-vertical" | "border-none"
     * @param sheet
     * @param borderInfo
     */
    private void rangeBorder(Sheet sheet, LuckySheetBorderInfoCellForImp borderInfo){
        List<LuckySheetCellRange> rangeList = borderInfo.getRange();
        String borderType = borderInfo.getBorderType();
        String style = borderInfo.getStyle();
        String color = borderInfo.getColor();
        for (LuckySheetCellRange range : rangeList) {

            switch(borderType) {
                case "border-left":rangeBorderLeft(sheet, style, color, range);break;
                case "border-right":rangeBorderRight(sheet, style, color, range);break;
                case "border-top":rangeBorderTop(sheet, style, color, range);break;
                case "border-bottom":rangeBorderBottom(sheet, style, color, range);break;
                case "border-all":rangeBorderAll(sheet,  style, color, range);break;
                case "border-outside":rangeBorderOutside(sheet,  style, color, range);break;
                case "border-inside":rangeBorderInside(sheet, style, color, range);break;
                case "border-horizontal":rangeBorderHorizontal(sheet, style, color, range);break;
                case "border-vertical":rangeBorderVertical(sheet, style, color, range);break;
                case "border-none":rangeBorderNone(sheet, style, color, range);break;
                default:break;
            }


        }

    }

    private void rangeBorderNone(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
                cellStyle.setBorderLeft(BorderStyle.NONE);
                cellStyle.setBorderRight(BorderStyle.NONE);
                cellStyle.setBorderTop(BorderStyle.NONE);
                cellStyle.setBorderBottom(BorderStyle.NONE);

                cellStyle.setTopBorderColor(convertBorderColor(color));
                cellStyle.setBottomBorderColor(convertBorderColor(color));
                cellStyle.setLeftBorderColor(convertBorderColor(color));
                cellStyle.setRightBorderColor(convertBorderColor(color));
            }
        }
    }

    private XSSFCellStyle getCellStyle(Sheet sheet, int row, int col) {
        XSSFCellStyle cellStyle;
        if(Objects.isNull(sheet.getRow(row))){
            sheet.createRow(row);
        }
        if(Objects.isNull(sheet.getRow(row).getCell(col))){
            sheet.getRow(row).createCell(col);
        }

        if(Objects.isNull(sheet.getRow(row).getCell(col).getCellStyle())){
            cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        }else{
            cellStyle = (XSSFCellStyle) sheet.getRow(row).getCell(col).getCellStyle();
            sheet.getRow(row).getCell(col).setCellStyle(cellStyle);
        }
        return cellStyle;
    }
    private void rangeBorderLeft(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int col = range.getColumnStart();

        for (int row = rowStart; row <= rowEnd; row++) {
            XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
            cellStyle.setBorderLeft(convertBorderType(Integer.parseInt(style)));
            cellStyle.setLeftBorderColor(convertBorderColor(color));

        }

    }

    private void rangeBorderRight(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int col = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
            cellStyle.setBorderRight(convertBorderType(Integer.parseInt(style)));
            cellStyle.setRightBorderColor(convertBorderColor(color));
        }
    }

    private void rangeBorderTop(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int row = range.getRowStart();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int col = colStart; col <= colEnd; col++) {
            XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
            cellStyle.setBorderTop(convertBorderType(Integer.parseInt(style)));
            cellStyle.setTopBorderColor(convertBorderColor(color));
        }
    }
    private void rangeBorderBottom(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int col = colStart; col <= colEnd; col++) {
            XSSFCellStyle cellStyle = getCellStyle(sheet, rowEnd, col);
            cellStyle.setBorderBottom(convertBorderType(Integer.parseInt(style)));
            cellStyle.setBottomBorderColor(convertBorderColor(color));
        }
    }

    private void rangeBorderAll(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
                cellStyle.setBorderTop(convertBorderType(Integer.parseInt(style)));
                cellStyle.setBorderBottom(convertBorderType(Integer.parseInt(style)));
                cellStyle.setBorderLeft(convertBorderType(Integer.parseInt(style)));
                cellStyle.setBorderRight(convertBorderType(Integer.parseInt(style)));

                cellStyle.setTopBorderColor(convertBorderColor(color));
                cellStyle.setBottomBorderColor(convertBorderColor(color));
                cellStyle.setLeftBorderColor(convertBorderColor(color));
                cellStyle.setRightBorderColor(convertBorderColor(color));

            }
        }
    }

//    "border-outside"
    private void rangeBorderOutside(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
                if( row == rowStart){
                    cellStyle.setBorderTop(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setTopBorderColor(convertBorderColor(color));
                }
                if( row == rowEnd){
                    cellStyle.setBorderBottom(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setBottomBorderColor(convertBorderColor(color));
                }
                if( col == colStart){
                    cellStyle.setBorderLeft(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setLeftBorderColor(convertBorderColor(color));
                }
                if( col == colEnd){
                    cellStyle.setBorderRight(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setRightBorderColor(convertBorderColor(color));
                }
            }
        }
    }
// "border-inside"
    private void rangeBorderInside(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
                if( row != rowStart && row != rowEnd){
                    cellStyle.setBorderTop(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setBorderBottom(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setTopBorderColor(convertBorderColor(color));
                    cellStyle.setBottomBorderColor(convertBorderColor(color));
                }
                if( col != colStart && col != colEnd){
                    cellStyle.setBorderLeft(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setBorderRight(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setLeftBorderColor(convertBorderColor(color));
                    cellStyle.setRightBorderColor(convertBorderColor(color));
                }
            }
        }
    }
//    "border-horizontal"
    private void rangeBorderHorizontal(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
                if( row != rowStart && row != rowEnd){
                    cellStyle.setBorderTop(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setBorderBottom(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setTopBorderColor(convertBorderColor(color));
                    cellStyle.setBottomBorderColor(convertBorderColor(color));
                }
            }
        }
    }
//    "border-vertical"
    private void rangeBorderVertical(Sheet sheet, String style, String color, LuckySheetCellRange range) {
        int rowStart = range.getRowStart();
        int rowEnd = range.getRowEnd();
        int colStart = range.getColumnStart();
        int colEnd = range.getColumnEnd();
        for (int row = rowStart; row <= rowEnd; row++) {
            for (int col = colStart; col <= colEnd; col++) {
                XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);
                if( col != colStart && col != colEnd){
                    cellStyle.setBorderLeft(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setBorderRight(convertBorderType(Integer.parseInt(style)));
                    cellStyle.setLeftBorderColor(convertBorderColor(color));
                    cellStyle.setRightBorderColor(convertBorderColor(color));
                }
            }
        }
    }

    private void cellBorder(Sheet sheet, LuckySheetBorderInfoCellForImp borderInfo) {
        LuckySheetBorderInfoCellValue value = borderInfo.getValue();
        int col = value.getColIndex();
        int row = value.getRowIndex();

        LuckySheetBorderInfoCellValueStyle left = value.getL();
        LuckySheetBorderInfoCellValueStyle right = value.getR();
        LuckySheetBorderInfoCellValueStyle top = value.getT();
        LuckySheetBorderInfoCellValueStyle bottom = value.getB();


        XSSFCellStyle cellStyle = getCellStyle(sheet, row, col);


        if(Objects.nonNull(left) ){
            cellStyle.setBorderLeft(convertBorderType(left.getStyle()));
            cellStyle.setLeftBorderColor(convertBorderColor(left.getColor()));
        }
        if(Objects.nonNull(right) ){
            cellStyle.setBorderRight(convertBorderType(right.getStyle()));
            cellStyle.setRightBorderColor(convertBorderColor(right.getColor()));
        }
        if(Objects.nonNull(top) ){
            cellStyle.setBorderTop(convertBorderType(top.getStyle()));
            cellStyle.setTopBorderColor(convertBorderColor(top.getColor()));
        }
        if(Objects.nonNull(bottom) ){
            cellStyle.setBorderBottom(convertBorderType(bottom.getStyle()));
            cellStyle.setBottomBorderColor(convertBorderColor(bottom.getColor()));
        }
    }


    private XSSFColor convertBorderColor(String rgbColor){
        return new XSSFColor(Util.toRGBBytes(rgbColor));
    }


    /**
     * style: 1 Thin | 2 Hair | 3 Dotted | 4 Dashed | 5 DashDot | 6 DashDotDot | 7 Double | 8 Medium |
     * 9 MediumDashed | 10 MediumDashDot | 11 MediumDashDotDot | 12 SlantedDashDot | 13 Thick
     * @param type
     * @return @BorderStyle
     */
    private BorderStyle convertBorderType(int type){

        BorderStyle borderStyle;
        switch (type){
            case 0: borderStyle = BorderStyle.NONE;break;
            case 1: borderStyle = BorderStyle.THIN;break;
            case 2: borderStyle = BorderStyle.HAIR;break;
            case 3: borderStyle = BorderStyle.DOTTED;break;
            case 4: borderStyle = BorderStyle.DASHED;break;
            case 5: borderStyle = BorderStyle.DASH_DOT;break;
            case 6: borderStyle = BorderStyle.DASH_DOT_DOT;break;
            case 7: borderStyle = BorderStyle.DOUBLE;break;
            case 8: borderStyle = BorderStyle.MEDIUM;break;
            case 9: borderStyle = BorderStyle.MEDIUM_DASHED;break;
            case 10: borderStyle = BorderStyle.MEDIUM_DASH_DOT;break;
            case 11: borderStyle = BorderStyle.MEDIUM_DASH_DOT_DOT;break;
            case 12: borderStyle = BorderStyle.SLANTED_DASH_DOT;break;
            case 13: borderStyle = BorderStyle.THICK;break;
            default: borderStyle = BorderStyle.NONE;break;
        }
        return borderStyle;
    }


    private void fillCellData(Sheet sheet, LuckySheetCellData[] celldata) {
        if (Objects.isNull(celldata)) {
            return;
        }
        CellConverter cellConverter = new CellConverter(this.workbook, sheet);
        for (LuckySheetCellData data : celldata) {
            // 设置单元格内容
            cellConverter.convert(data);
        }
    }

//    getptToPxRatioByDPI():number{
//        return 72/96;
//    }
    private Float getptToPxRatioByDPI(){
        return 72f/96;
    }

    // todo 控制图片位置要用到
    private int getPxByEMUs(Integer emus){
        if(emus==null){
            return 0;
        }
        Float inch = emus.floatValue()/914400;
        Float pt = inch*72;
        return (int) (pt/getptToPxRatioByDPI());
    }
    private Integer getEMUsByPx(int px) {
        // 将像素转换为点
        float pt = px * getptToPxRatioByDPI();
        // 将点转换为英寸
        float inch = pt / 72;

//        float inch = px * 96
        // 将英寸转换为 EMUs
        Integer emus = (int) (inch * 914400);
        return emus;
    }


    private int calcColWidth(int pixelWidth) {
         return (int) (Units.pixelToEMU(pixelWidth)*256d/Units.EMU_PER_CHARACTER);
    }
    private short calcRowHeight(int pixelHeight) {
        return (short) (Units.pixelToPoints(pixelHeight)*20);
    }

    private void setColumnWidth(Sheet sheet, Map<String, Integer> columnlen) {
        if (Objects.isNull(columnlen)) {
            return;
        }
        for (Map.Entry<String, Integer> entry : columnlen.entrySet()) {
            int colWidth = calcColWidth(entry.getValue());
            sheet.setColumnWidth(Integer.parseInt(entry.getKey()), colWidth);
        }
    }

    private void setRowHeight(Sheet sheet, Map<String, Integer> rowlen) {
        if (Objects.isNull(rowlen)) {
            return;
        }
        for (Map.Entry<String, Integer> entry : rowlen.entrySet()) {
            short rowHeight = calcRowHeight(entry.getValue());
            Row row = sheet.getRow(Integer.parseInt(entry.getKey()));
            row = Objects.isNull(row) ? sheet.createRow(Integer.parseInt(entry.getKey())) : row;
            row.setHeight(rowHeight);
        }
    }
    private void setRowHidden(Sheet sheet, Map<String, Integer> rowhidden) {
        if (Objects.isNull(rowhidden)) {
            return;
        }
        for (Map.Entry<String, Integer> entry : rowhidden.entrySet()) {
            Row row = sheet.getRow(Integer.parseInt(entry.getKey()));
            row = Objects.isNull(row) ? sheet.createRow(Integer.parseInt(entry.getKey())) : row;
            row.setZeroHeight(entry.getValue() == 0);
        }
    }
    private void setColumnHidden(Sheet sheet, Map<String, Integer> colhidden) {
        if (Objects.isNull(colhidden)) {
            return;
        }
        for (Map.Entry<String, Integer> entry : colhidden.entrySet()) {
            sheet.setColumnHidden(Integer.parseInt(entry.getKey()), entry.getValue() == 0);
        }
    }

    private static void mergeCell(Sheet Sheet, Map<String, LuckySheetConfigMerge> merge) {
        if (Objects.isNull(merge)) {
            return;
        }
        for (Map.Entry<String, LuckySheetConfigMerge> entry : merge.entrySet()) {
            LuckySheetConfigMerge value = entry.getValue();
            int firstRow = value.getR();
            int lastRow = value.getR() + value.getRs() - 1;
            int firstCol = value.getC();
            int lastCol = value.getC() + value.getCs() - 1;
            Sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        }
    }
}
