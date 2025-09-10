package com.tugalsan.api.file.xlsx.server;

import com.tugalsan.api.function.client.maythrowexceptions.unchecked.TGS_FuncMTUEffectivelyFinal;
import com.tugalsan.api.charset.client.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.util.*;
import org.apache.poi.xssf.usermodel.*;
import com.tugalsan.api.list.client.*;
import com.tugalsan.api.log.server.*;
import com.tugalsan.api.file.server.*;
import com.tugalsan.api.function.client.maythrowexceptions.checked.TGS_FuncMTCUtils;
import com.tugalsan.api.function.client.maythrowexceptions.unchecked.TGS_FuncMTUUtils;
import com.tugalsan.api.string.client.*;

import java.util.stream.*;

public class TS_FileXlsxUtils implements Closeable {

    final private static TS_Log d = TS_Log.of(TS_FileXlsxUtils.class);

    private static int CELL_TEXT_MAX_CHAR_SIZE() {
        return 32767;
    }

    private static int DEFAULT_COLUMNWIDTH() {
        return 1;
    }//DONT CHANGE!! //EXCEL MAX 255 char

    private static double DEFAULT_COLUMNWIDTH_INPX() {
        return 7.0017f;
    }//DONT CHANGE!! //EXCEL MAX 255 char

    private static String DEFAULT_FONTNAME() {
        return "Calibri";
    }//DONT CHANGE!

    private static String DEFAULT_AUTHOR() {
        return "(mesametal.com & mebosa.com) by Tuğalsan Karabacak";
    }//DONT CHANGE!

    private static String DEFAULT_SHEETNAME() {
        return "Sayfa1";
    }//DONT CHANGE!

    private Workbook workbook = null;
    private CreationHelper creationHelper = null;
    private Sheet sheet = null;
    private final Path filePath;
    private final CellStyle cellStyle_LeftTopBordered;
    private final CellStyle cellStyle_RightTopBordered;
    private final CellStyle cellStyle_CenterTopBordered;
    private final CellStyle cellStyle_LeftTop;
    private final CellStyle cellStyle_RightTop;
    private final CellStyle cellStyle_CenterTop;

    public Path getFile() {
        return filePath;
    }

    public final void setAuthor(CharSequence author) {
        switch (workbook) {
            case XSSFWorkbook wb ->
                wb.getProperties().getCoreProperties().setCreator(author.toString());
            case HSSFWorkbook wb -> {
                wb.createInformationProperties();
                wb.getSummaryInformation().setAuthor(author.toString());
            }
            default -> {
            }
        }
    }

    public int getMaxPageColumnCount() {
        return getMaxPageColumnCount(sheet.getPrintSetup().getLandscape(), sheet.getPrintSetup().getPaperSize());
    }

    public int getMaxPageColumnCount(boolean landscape, int pageSizeAX) {
        return switch (pageSizeAX) {
            case HSSFPrintSetup.A3_PAPERSIZE ->
                landscape ? 88 : 59;
            case HSSFPrintSetup.A5_PAPERSIZE ->
                landscape ? 39 : 25;
            default ->
                landscape ? 59 : 39;
        }; //A4
    }

    public final void setPageSize(boolean landscape, int pageSizeAX) {
        sheet.getPrintSetup().setLandscape(landscape);
        pageSizeAX = pageSizeAX < 3 ? 3 : pageSizeAX;
        pageSizeAX = pageSizeAX > 5 ? 5 : pageSizeAX;
        switch (pageSizeAX) {
            case 3 -> {
                sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.A3_PAPERSIZE);
                d.ci("setPageSize", "A", 3);
            }
            case 4 -> {
                sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.A4_PAPERSIZE);
                d.ci("setPageSize", "A", 4);
            }
            case 5 -> {
                sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.A5_PAPERSIZE);
                d.ci("setPageSize", "A", 5);
            }
            default -> {
            }
        }
    }

    private int getCharCount(List<String> texts) {
        var charCount = 0;
        charCount = texts.stream().map(t -> t.length()).reduce(charCount, Integer::sum);
        return charCount > CELL_TEXT_MAX_CHAR_SIZE() ? CELL_TEXT_MAX_CHAR_SIZE() : charCount;
    }

    private int getMaxFontHeight(List<Font> fonts) {
        var maxFontSize = fonts.stream().mapToInt(f -> f.getFontHeightInPoints()).max().orElse(1);
        return maxFontSize < 1 ? 1 : maxFontSize;
    }

    @Deprecated
    public void setRowHeightAuto(int rowIdx) {
        getRow(rowIdx).setHeight((short) -1);
//        getRow(rowIdx).setHeight((short) 0);
    }

    public int calculateCellHeight(List<String> texts, List<Font> fonts, int colSpan) {
        var maxFontSize = getMaxFontHeight(fonts);
        var singleRowHeight = calculateSingleRow_RowHeight(maxFontSize);

        var charCount = getCharCount(texts);
        var cellSingleRowCapacity = calculateSingleRow_CellCharCapacity(colSpan, colSpan);
        var rowCount = (charCount / cellSingleRowCapacity) + 1;
        d.ci("calculateCellHeight", "maxFontSize", maxFontSize, "singleRowHeight", singleRowHeight,
                "charCount", charCount, "cellSingleRowCapacity", cellSingleRowCapacity, "rowCount", rowCount);
        return singleRowHeight * rowCount;
    }

    private int calculateSingleRow_RowHeight(int fontSize) {
        fontSize = fontSize < 1 ? 1 : fontSize;
        fontSize = fontSize > 72 ? 72 : fontSize;
        var rowSize = fontSize < 42 ? (1.3036d * fontSize + 1) : (1.4243d * fontSize - 3.984d);
        return Math.min(409, (int) Math.ceil(rowSize));
    }

    private int calculateSingleRow_CellCharCapacity(int fontSize, int colSpan) {
        fontSize = fontSize < 1 ? 1 : fontSize;
        fontSize = fontSize > 72 ? 72 : fontSize;
        colSpan = colSpan < 1 ? 1 : colSpan;

        //charSize of char 'A' for colspan 39 A4 PORT
        double rowCharSizeA4Port;
        if (fontSize > 49) {
            rowCharSizeA4Port = -0.2332d * fontSize + 27.399d;
        } else if (fontSize > 34) {
            rowCharSizeA4Port = -0.4786d * fontSize + 39.433d;
        } else if (fontSize > 22) {
            rowCharSizeA4Port = -1.0245d * fontSize + 58.114d;
        } else if (fontSize > 15) {
            rowCharSizeA4Port = -2.2857d * fontSize + 86.857d;
        } else if (fontSize > 4) {
            rowCharSizeA4Port = -21.2d * fontSize + 268.3d;
        } else if (fontSize > 2) {
            rowCharSizeA4Port = -75d * fontSize + 509d;
        } else {
            rowCharSizeA4Port = -399d * fontSize + 1196d;
        }
        rowCharSizeA4Port *= (27.108d * colSpan + 18.973d) / (27.108d * getMaxPageColumnCount(false, 4) + 18.973d);
        return (int) Math.floor(rowCharSizeA4Port);
    }

    public TS_FileXlsxUtils(Path destXLSX) {
        this.filePath = destXLSX;

        workbook = new XSSFWorkbook();//xls: HSSFWorkbook
        setAuthor(DEFAULT_AUTHOR());
        workbook.setForceFormulaRecalculation(true);
        creationHelper = workbook.getCreationHelper();
        sheet = workbook.createSheet(DEFAULT_SHEETNAME());
        sheet.setDefaultColumnWidth(DEFAULT_COLUMNWIDTH());
        setPageSize(false, 4);
        {
            cellStyle_LeftTopBordered = workbook.createCellStyle();
            setColorBackgroundSolid(cellStyle_LeftTopBordered, getIndexedColorWHITE());
            cellStyle_LeftTopBordered.setWrapText(true);
            cellStyle_LeftTopBordered.setAlignment(HorizontalAlignment.LEFT);
            cellStyle_LeftTopBordered.setVerticalAlignment(VerticalAlignment.TOP);
        }
        {
            cellStyle_RightTopBordered = workbook.createCellStyle();
            setColorBackgroundSolid(cellStyle_RightTopBordered, getIndexedColorWHITE());
            cellStyle_RightTopBordered.setWrapText(true);
            cellStyle_RightTopBordered.setAlignment(HorizontalAlignment.RIGHT);
            cellStyle_RightTopBordered.setVerticalAlignment(VerticalAlignment.TOP);
        }
        {
            cellStyle_CenterTopBordered = workbook.createCellStyle();
            setColorBackgroundSolid(cellStyle_CenterTopBordered, getIndexedColorWHITE());
            cellStyle_CenterTopBordered.setWrapText(true);
            cellStyle_CenterTopBordered.setAlignment(HorizontalAlignment.CENTER);
            cellStyle_CenterTopBordered.setVerticalAlignment(VerticalAlignment.TOP);
        }
        {
            cellStyle_LeftTop = workbook.createCellStyle();
            setColorBackgroundSolid(cellStyle_LeftTop, getIndexedColorWHITE());
            cellStyle_LeftTop.setWrapText(true);
            cellStyle_LeftTop.setAlignment(HorizontalAlignment.LEFT);
            cellStyle_LeftTop.setVerticalAlignment(VerticalAlignment.TOP);
        }
        {
            cellStyle_RightTop = workbook.createCellStyle();
            setColorBackgroundSolid(cellStyle_RightTop, getIndexedColorWHITE());
            cellStyle_RightTop.setWrapText(true);
            cellStyle_RightTop.setAlignment(HorizontalAlignment.RIGHT);
            cellStyle_RightTop.setVerticalAlignment(VerticalAlignment.TOP);
        }
        {
            cellStyle_CenterTop = workbook.createCellStyle();
            setColorBackgroundSolid(cellStyle_CenterTop, getIndexedColorWHITE());
            cellStyle_CenterTop.setWrapText(true);
            cellStyle_CenterTop.setAlignment(HorizontalAlignment.CENTER);
            cellStyle_CenterTop.setVerticalAlignment(VerticalAlignment.TOP);
        }
    }

    public void setCellStyle(Cell cell, CellStyle cs) {
        cell.setCellStyle(cs);
    }

    public CellStyle createCellStyle(int allign_center1_right2_defaultLeft, boolean isBordered) {
        var cs = workbook.createCellStyle();
        switch (allign_center1_right2_defaultLeft) {
            case 1 ->
                cs.cloneStyleFrom(isBordered ? cellStyle_CenterTopBordered : cellStyle_CenterTop);
            case 2 ->
                cs.cloneStyleFrom(isBordered ? cellStyle_RightTopBordered : cellStyle_RightTop);
            default ->
                cs.cloneStyleFrom(isBordered ? cellStyle_LeftTopBordered : cellStyle_LeftTop);
        }
        return cs;
    }

    public Font getDefaultFont() {
        if (DEFAULY_FONT == null) {
            DEFAULY_FONT = createFont(false, false, false, 12, IndexedColors.BLACK.getIndex());
        }
        return DEFAULY_FONT;
    }
    private Font DEFAULY_FONT = null;

    public CellStyle getDefaultCellStyle() {
        if (DEFAULY_CELL_STYLE == null) {
            DEFAULY_CELL_STYLE = createCellStyle(0, true);
        }
        return DEFAULY_CELL_STYLE;
    }
    private CellStyle DEFAULY_CELL_STYLE = null;

    public Row getRow(int rowNumber_starts0) {
        var row = sheet.getRow(rowNumber_starts0);
        if (row == null) {
            row = sheet.createRow((short) rowNumber_starts0);
        }
        return row;
    }

    public Cell getCell(int rowNumber_starts0, int cellColumn_starts0) {
        var row = getRow(rowNumber_starts0);
        var cell = row.getCell(cellColumn_starts0);
        if (cell == null) {
            cell = row.createCell((short) cellColumn_starts0);
        }
        return cell;
    }

    public RichTextString createRichText(CharSequence text, Font font) {
        var textStr = text.toString();
        textStr = textStr.length() > CELL_TEXT_MAX_CHAR_SIZE() ? TGS_StringUtils.cmn().concat(textStr.substring(0, CELL_TEXT_MAX_CHAR_SIZE() - 3), "...") : textStr;
        var rts = creationHelper.createRichTextString(textStr);
        rts.applyFont(font);
        return rts;
    }

    public RichTextString createRichText(List<String> texts, List<Font> fonts) {
        TGS_ListCleanUtils.cleanNulls(texts, i -> fonts.remove(i));
        var sb = new StringBuilder();
        texts.stream().forEachOrdered(s -> sb.append(s));
        var text = sb.toString();
        text = text.length() > CELL_TEXT_MAX_CHAR_SIZE() ? text.substring(0, CELL_TEXT_MAX_CHAR_SIZE()) : text;
        var rts = creationHelper.createRichTextString(text);

        var offset = 0;
        int length;
        for (var i = 0; i < texts.size(); i++) {//style
            length = texts.get(i).length();
            if (offset + length > CELL_TEXT_MAX_CHAR_SIZE()) {
                break;
            }
            rts.applyFont(offset, offset + length, fonts.get(i));
            offset += length;
        }
        return rts;
    }

    public void setRowHeight(int rowIdx, float row_height_default_is_15) {
        row_height_default_is_15 = row_height_default_is_15 > 409 ? 409 : row_height_default_is_15;
        getRow(rowIdx).setHeightInPoints(row_height_default_is_15);
    }

    public void setCellRichText(Cell cell, RichTextString rts) {
        cell.setCellValue(rts);
    }

    public void setCellHyperlink(Cell cell, CharSequence strhttp, CharSequence strName) {
        cell.setCellFormula(TGS_StringUtils.cmn().concat("HYPERLINK(\"", strhttp, "\",\"" + strName, "\")"));
    }

    public void setFontSize(Font font, int height) {
        font.setFontHeightInPoints((short) height);
    }

    public int getFontSize(Font font) {
        return (int) font.getFontHeightInPoints();
    }

    public void setFontUnderline(Font font, boolean isUnderline) {
        font.setUnderline(isUnderline ? Font.U_SINGLE : Font.U_NONE);
    }

    public void setFontColor(Font font, short indexedColor) {
        font.setColor(indexedColor);
    }

    public static short getIndexedColorWHITE() {
        return IndexedColors.WHITE.getIndex();
    }

    public static short getIndexedColorBLUE() {
        return IndexedColors.BLUE.getIndex();
    }

    public static short getIndexedColorRED() {
        return IndexedColors.RED.getIndex();
    }

    public static short getIndexedColorBLACK() {
        return IndexedColors.BLACK.getIndex();
    }

    public static short getIndexedColorGREEN() {
        return IndexedColors.GREEN.getIndex();
    }

    public Font createFont(boolean isBold, boolean isItalic, boolean isUnderline, Integer height, Short indexedColor) {
        var font = workbook.createFont();
        font.setFontName(DEFAULT_FONTNAME());
        font.setBold(isBold);
        font.setItalic(isItalic);
        font.setUnderline(isUnderline ? Font.U_SINGLE : Font.U_NONE);
        if (height != null) {
            height = height < 1 ? 1 : height;
            height = height > 72 ? 72 : height;
            font.setFontHeightInPoints(height.shortValue());
        }
        font.setColor(indexedColor == null ? getIndexedColorBLACK() : indexedColor);
        return font;
    }

    public Font createFont(boolean isBold, boolean isItalic, boolean isUnderline) {
        return createFont(isBold, isItalic, isUnderline, null, null);
    }

    public CellRangeAddress createMergedCell(int rowFrom, int rowTo, int colFrom, int colTo, boolean ignoreExceptions) {
        return TGS_FuncMTCUtils.call(() -> {
            d.ci("createMergedCell rf:" + rowFrom + ", rt:" + rowTo + ", cf:" + colFrom + ", ct:" + colTo);
            var rangeAddress = new CellRangeAddress(rowFrom, rowTo, colFrom, colTo);
            sheet.addMergedRegion(rangeAddress);
            return rangeAddress;
        }, e -> {
            if (ignoreExceptions) {
                return null;
            }
            return TGS_FuncMTUUtils.thrw(e);
        });
    }

    final public void setColorBackgroundSolid(CellStyle cellStyle, short indexedColor) {
        cellStyle.setFillForegroundColor(indexedColor);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public void setColorBackgroundSolid(Cell cell, short indexedColor) {
        var cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cellStyle = getDefaultCellStyle();//cell.getSheet().getWorkbook().createCellStyle();
        }
        setColorBackgroundSolid(cellStyle, indexedColor);
        cell.setCellStyle(cellStyle);
    }

    public void setColorBackgroundBigSpots(CellStyle cellStyle, short indexedColor1, short indexedColor2) {
        cellStyle.setFillBackgroundColor(indexedColor1);
        cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyle.setFillForegroundColor(indexedColor2);
    }

    public void setColorBackgroundBigSpots(Cell cell, short indexedColor1, short indexedColor2) {
        var cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cellStyle = getDefaultCellStyle();//cell.getSheet().getWorkbook().createCellStyle();
        }
        setColorBackgroundBigSpots(cellStyle, indexedColor1, indexedColor2);
        cell.setCellStyle(cellStyle);
    }

    public void setBordersToMergedCell(CellRangeAddress rangeAddress, boolean ignoreExceptions) {
        TGS_FuncMTCUtils.run(() -> {
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.MEDIUM, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.MEDIUM, rangeAddress, sheet);
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, rangeAddress, sheet);
        }, e -> {
            if (ignoreExceptions) {
                return;
            }
            TGS_FuncMTUUtils.thrw(e);
        });
    }

    @Override
    public void close() {
        close(false);
    }

    public void close(boolean ignoreExceptions) {
        TGS_FuncMTCUtils.run(() -> {
            try (var fileOut = new FileOutputStream(filePath.toFile())) {
                workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
                workbook.write(fileOut);
            }
        }, e -> {
            if (ignoreExceptions) {
                return;
            }
            d.ce("close", e);
            TGS_FuncMTUUtils.thrw(e);
        });
    }

    public void addImage(CharSequence imgFile, int rowIdx, int colIdx, int colspan) {
        TGS_FuncMTCUtils.run(() -> {
            var imgFileStr = imgFile.toString();
            var imgFileStrLc = TGS_CharSetCast.current().toLowerCase(imgFileStr);
            var format = TGS_FuncMTUEffectivelyFinal.ofInt()
                    .anointAndCoronateIf(val -> imgFileStrLc.endsWith(".emf"), val -> Workbook.PICTURE_TYPE_EMF)
                    .anointAndCoronateIf(val -> imgFileStrLc.endsWith(".wmf"), val -> Workbook.PICTURE_TYPE_WMF)
                    .anointAndCoronateIf(val -> imgFileStrLc.endsWith(".png"), val -> Workbook.PICTURE_TYPE_PNG)
                    .anointAndCoronateIf(val -> imgFileStrLc.endsWith(".jpeg") || imgFileStrLc.endsWith(".jpg"), val -> Workbook.PICTURE_TYPE_JPEG)
                    .anointAndCoronateIf(val -> imgFileStrLc.endsWith(".pict") || imgFileStrLc.endsWith(".pıct"), val -> Workbook.PICTURE_TYPE_PICT)
                    .anointAndCoronateIf(val -> imgFileStrLc.endsWith(".dib") || imgFileStrLc.endsWith(".dıb"), val -> Workbook.PICTURE_TYPE_DIB)
                    .coronate();
            if (format == null) {
                d.ce("addImage.Unsupported picture: " + imgFileStr + ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
                return;
            }
//        r.setText(imgFile);
//        r.addBreak();
            d.ci("addImage.INFO: TK_XLSXFile.run.addPicture.BEGIN...");
            var pictureIdx = workbook.addPicture(TS_FileUtils.read(Path.of(imgFileStr)), format);

            var helper = workbook.getCreationHelper();
            var anchor = helper.createClientAnchor();
            anchor.setCol1(colIdx);
            anchor.setRow1(rowIdx);

            var drawing = sheet.createDrawingPatriarch();
            var pict = drawing.createPicture(anchor, pictureIdx);
            pict.resize();//resize to original size
            d.ci("addImage: w:" + pict.getImageDimension().getWidth() + ", h:" + pict.getImageDimension().getHeight() + ", l:" + imgFileStr);

            //get calculate width center 
            var pictWidthPx = pict.getImageDimension().width;
            d.ci("addImage.pictWidthPx:" + pictWidthPx);
            var cellWidthPx = IntStream.range(0, colspan)
                    .mapToLong(ci -> Math.round(DEFAULT_COLUMNWIDTH_INPX()))
                    .sum();
            d.ci("addImage.cellWidthPx:" + cellWidthPx);
            var centerPosPx = Math.round(cellWidthPx / 2d - pictWidthPx / 2d);
            d.ci("addImage.centerPosPx:" + centerPosPx);
            d.ci("addImage.pictWidthPx:" + pictWidthPx + ", cellWidthPx:" + cellWidthPx + ", centerPosPx:" + centerPosPx);

            //determine the new first anchor column dependent of the center position 
            //and the remaining pixels as Dx
            var anchorCol1 = 0;
            for (var ci = 0; ci < colspan; ci++) {
                if (Math.round(DEFAULT_COLUMNWIDTH_INPX()) < centerPosPx) {
                    centerPosPx -= Math.round(DEFAULT_COLUMNWIDTH_INPX());
                    anchorCol1 = ci + 1;
                } else {
                    break;
                }
            }
            d.ci("addImage.anchorCol1:" + anchorCol1);
            d.ci("addImage.centerPosPx:" + centerPosPx);

            anchor.setCol1(anchorCol1 < 0 ? 0 : anchorCol1);//set the new upper left anchor position
            anchor.setDx1((int) (centerPosPx * Units.EMU_PER_PIXEL));//set the remaining pixels up to the center position as Dx in unit EMU
            pict.resize(); //resize the pictutre to original size again; this will determine the new bottom rigth anchor position
            d.ci("addImage.END");
        });
    }

    public void setCell(int ri, int ci, CellStyle cs, List<Font> font, List<String> text) {
        var xlsxCell = getCell(ri, 0);
        xlsxCell.setCellValue(createRichText(text, font));
        xlsxCell.setCellStyle(cs);
    }

    public void setCell(int ri, int ci, CellStyle cs, Font font, String text) {
        var xlsxCell = getCell(ri, 0);
        xlsxCell.setCellValue(createRichText(text, font));
        xlsxCell.setCellStyle(cs);
    }

    public Workbook getWorkBook() {
        return workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }
}
