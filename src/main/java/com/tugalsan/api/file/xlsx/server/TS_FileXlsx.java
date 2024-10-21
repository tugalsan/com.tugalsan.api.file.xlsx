package com.tugalsan.api.file.xlsx.server;

import com.tugalsan.api.function.client.TGS_Func_In1;
import com.tugalsan.api.file.common.server.TS_FileCommonAbstract;
import com.tugalsan.api.file.common.server.TS_FileCommonFontTags;
import com.tugalsan.api.file.common.server.TS_FileCommonConfig;
import com.tugalsan.api.string.client.*;
import com.tugalsan.api.file.server.*;
import com.tugalsan.api.cast.client.*;
import java.awt.image.*;
import java.nio.file.*;
import java.util.*;
import java.util.stream.*;
import org.apache.poi.ss.usermodel.*;
import com.tugalsan.api.math.client.*;
import com.tugalsan.api.log.server.*;
import com.tugalsan.api.list.client.*;

import com.tugalsan.api.union.client.TGS_UnionExcuseVoid;
import com.tugalsan.api.unsafe.client.*;
import com.tugalsan.api.url.client.*;

public class TS_FileXlsx extends TS_FileCommonAbstract {

    final private static TS_Log d = TS_Log.of(TS_FileXlsx.class);

    private static int FONT_HEIGHT_OFFSET() {
        return -2;
    }

    private TS_FileXlsxUtils xlsx;

    private final List doc = TGS_ListUtils.of();
    private TGS_ListTable table;
    private int[] table_relColSizes = null;
    private int currentRowIndex = 0;
    private int currentColXLSXIndex;
    private int currentColRelIndex = 0;
    private TS_MIFXLSX_RichCell lastCell = null;
    private final String CELL_Occupied = "CELL_FULL";
    private final String CELL_Vacant = "CELL_EMPTY";

    private TS_FileCommonConfig fileCommonConfig;
    private boolean landscape = false;
    private int pageSizeAX = 4;

    private TS_FileXlsx(boolean enabled, Path localFile, TGS_Url remoteFile) {
        super(enabled, localFile, remoteFile);
    }

    public static void use(boolean enabled, TS_FileCommonConfig fileCommonConfig, Path localFile, TGS_Url remoteFile, TGS_Func_In1<TS_FileXlsx> xlsx) {
        var instance = new TS_FileXlsx(enabled, localFile, remoteFile);
        try {
            instance.use_init(fileCommonConfig);
            xlsx.run(instance);
        } catch (Exception e) {
            instance.saveFile(e.getMessage());
            throw e;
        } finally {
            instance.saveFile(null);
        }

    }

    private void use_init(TS_FileCommonConfig fileCommonConfig) {
        this.fileCommonConfig = fileCommonConfig;
        if (isClosed()) {
            return;
        }
        xlsx = new TS_FileXlsxUtils(localFile);
        createCurrentFont();
    }

    @Override
    public boolean addLineBreak() {
        if (isClosed()) {
            return true;
        }
        return addText("\n");
    }

    @Override
    public boolean setFontStyle() {
        if (isClosed()) {
            return true;
        }
        return createCurrentFont();
    }

    @Override
    public boolean setFontHeight() {
        if (isClosed()) {
            return true;
        }
        return createCurrentFont();
    }

    @Override
    public boolean setFontColor() {
        if (isClosed()) {
            return true;
        }
        return createCurrentFont();
    }

    private boolean createCurrentFont() {
        if (isClosed()) {
            return true;
        }
        var fh = fileCommonConfig.fontHeight + FONT_HEIGHT_OFFSET();
        if (fh < 1) {
            fh = 1;
        }
        currentFont = xlsx.createFont(fileCommonConfig.fontBold, fileCommonConfig.fontItalic,
                fileCommonConfig.fontUnderlined, fh, getFontColor());
        var k = 0.8f;
        var fh_half = (int) (fh * k);
        if (fh_half < 1) {
            fh_half = 1;
        }
        currentFont_half = xlsx.createFont(fileCommonConfig.fontBold, fileCommonConfig.fontItalic,
                fileCommonConfig.fontUnderlined, fh_half, getFontColor());
        return true;
    }
    private Font currentFont;
    private Font currentFont_half;

    private short getFontColor() {
        if (isClosed()) {
            return IndexedColors.BLACK.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_BLUE())) {
            return IndexedColors.BLUE.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_CYAN())) {
            return IndexedColors.CORAL.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_DARK_GRAY())) {
            return IndexedColors.GREY_80_PERCENT.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_GRAY())) {
            return IndexedColors.GREY_40_PERCENT.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_GREEN())) {
            return IndexedColors.GREEN.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_LIGHT_GRAY())) {
            return IndexedColors.GREY_25_PERCENT.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_MAGENTA())) {
            return IndexedColors.MAROON.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_ORANGE())) {
            return IndexedColors.ORANGE.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_PINK())) {
            return IndexedColors.PINK.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_RED())) {
            return IndexedColors.RED.getIndex();
        }
        if (Objects.equals(fileCommonConfig.fontColor, TS_FileCommonFontTags.CODE_TOKEN_FONT_COLOR_YELLOW())) {
            return IndexedColors.YELLOW.getIndex();
        }
        return IndexedColors.BLACK.getIndex();
    }

    @Override
    public boolean addText(String text) {
        if (isClosed()) {
            return true;
        }
        if (lastCell == null) {
            d.ce("addText. why lastCell not exists!");
            return false;
        }
//        d.infoEnable = true;
        d.ci("addText", text);
        var lines = TGS_StringUtils.jre().toList(text, "\n");
        IntStream.range(0, lines.size()).forEachOrdered(i -> {
            var line = lines.get(i);
            if (!line.isEmpty()) {
                if (!TGS_StringDouble.may(text)) {
                    d.ci("addText", "line", "mayNot", line);
                    lastCell.texts.add(line);
                    lastCell.fonts.add(currentFont);
                } else {
                    var tags = TGS_StringUtils.jre().toList_spc(line);
                    IntStream.range(0, tags.size()).forEachOrdered(j -> {
                        var tag = tags.get(j);
                        var dbl = TGS_StringDouble.of(text);
                        if (dbl.isExcuse()) {
                            lastCell.texts.add(tag);
                            lastCell.fonts.add(currentFont);
                        } else {
                            lastCell.texts.add(String.valueOf(dbl.value().left));
                            lastCell.fonts.add(currentFont);
                            lastCell.texts.add(String.valueOf(dbl.value().dim()) + String.valueOf(dbl.value().right));
                            lastCell.fonts.add(currentFont_half);
                        }
                        if (tags.size() - 1 != j) {
                            lastCell.texts.add(" ");
                            lastCell.fonts.add(currentFont);
                        }
                    });
                }
            }
            if (i != lines.size() - 1 || text.endsWith("\n")) {
                lastCell.texts.add("\n");
                lastCell.fonts.add(currentFont);
            }
        });
        return true;
    }

    @Override
    public boolean createNewPage(int pageSizeAX, boolean landscape, Integer marginLeft, Integer marginRight, Integer marginTop, Integer marginBottom) {
        if (isClosed()) {
            return true;
        }
        if (table != null) {
            d.ce("createNewPage. why table exists!");
            return false;
        }
        if (firstPageTriggered) {
            currentRowIndex++;
            beginText(0);
            addText("--------------------------------------------------------------------------");
            endText();
        } else {
            firstPageTriggered = true;
            if (!pageSizeTriggered) {
                this.landscape = landscape;
                this.pageSizeAX = pageSizeAX;
                pageSizeTriggered = true;
            }
        }
        return true;
    }
    private boolean firstPageTriggered = false;
    private boolean pageSizeTriggered = false;

    @Override
    public boolean addImage(BufferedImage pstImage, Path pstImageLoc, boolean textWrap, int left0_center1_right2, long imageCounter) {
        if (isClosed()) {
            return true;
        }
        d.ci("addImage", "init", "imageLoc", pstImageLoc);
        if (lastCell == null) {
            d.ce("addImage. why lastCell not exists!");
            return false;
        }
        lastCell.imgFiles.add(pstImageLoc);
        d.ci("addImage", "fin");
        return true;
    }

    @Override
    public boolean saveFile(String errorSource) {
        if (isClosed()) {
            return true;
        }
        setClosed();
        d.ci("saveFile.XLSX->");
        if (xlsx == null) {
            d.ci("XLSX File is null");
        } else {
            var u = saveFile_close();
            if (u.isExcuse()) {
                d.ce("compileFile.ERROR: MIFXLSX.close -> " + u.excuse().getMessage());
                u.excuse().printStackTrace();
            }
            if (TS_FileUtils.isExistFile(xlsx.getFile())) {
                d.ci("saveFile.FIX: XLSX File save", xlsx.getFile(), "successfull");
            } else {
                d.ce("saveFile.FIX: XLSX File save", xlsx.getFile(), "failed");
            }
        }
        return errorSource == null;
    }

    @Override
    public boolean beginTable(int[] relColSizes) {
        if (isClosed()) {
            return true;
        }
        if (table != null) {
            d.ce("beginTable.ERROR: MIFXLSX.beginTableCell -> why table exists!");
            return false;
        }
        this.table_relColSizes = TGS_MathUtils.convertWeighted(relColSizes, 1, xlsx.getMaxPageColumnCount());
        d.ci("beginTable.PST relColSizes: " + TGS_StringUtils.cmn().toString(table_relColSizes, ", "));
        table = TGS_ListTable.of(false);
        currentRowIndex = 0;
        currentColXLSXIndex = 0;
        currentColRelIndex = 0;
        d.ci("beginTable.currentColRelIndex:", currentColRelIndex);
        d.ci("beginTable.currentColXLSXIndex:", currentColXLSXIndex);
        if (table.getColumnSize(currentRowIndex) < relColSizes.length) {
            for (var ci = 0; ci < relColSizes.length; ci++) {
                if (table.getValueAsObject(currentRowIndex, ci) == null) {
                    table.setValue(currentRowIndex, ci, CELL_Vacant);
                }
            }
        }
        doc.add(table);
        return true;
    }

    @Override
    public boolean endTable() {
        if (isClosed()) {
            return true;
        }
        this.table_relColSizes = null;
        table = null;
        return true;
    }

    private int findMeAndEmptyCell(int rowIdx) {
        var emptyCellIdx = table.getColumnSize(rowIdx);
        if (isClosed()) {
            return emptyCellIdx;
        }
        for (var ci = 0; ci < table.getColumnSize(rowIdx); ci++) {
            var o = table.getValueAsObject(rowIdx, ci);
            if (o instanceof String ins) {
                if (CELL_Vacant.equals(ins)) {
                    emptyCellIdx = ci;
                    break;
                }
            }
        }
        return emptyCellIdx;
    }

    @Override
    public boolean beginTableCell(int rowSpan, int colSpan, Integer cellHeight) {
        if (isClosed()) {
            return true;
        }
        while (true) {
            currentColXLSXIndex = findMeAndEmptyCell(currentRowIndex);
            d.ci("beginTableCell.currentRowIndex0:", currentRowIndex);
            if (currentColXLSXIndex >= xlsx.getMaxPageColumnCount()) {
                currentRowIndex++;
                d.ci("beginTableCell.currentRowIndex1:", currentRowIndex);
            } else {
                break;
            }
        }
        currentColXLSXIndex = findMeAndEmptyCell(currentRowIndex);
        d.ci("beginTableCell.currentColXLSXIndex2:", currentColXLSXIndex);
        currentColRelIndex = calcColRelIndex(currentRowIndex);
        d.ci("beginTableCell.currentColRelIndex2:", currentColRelIndex);

        var colSpanSize = 0;
        for (var ci = currentColRelIndex; ci < currentColRelIndex + colSpan; ci++) {
            if (ci >= table_relColSizes.length) {
                colSpanSize++;
            } else {
                colSpanSize += table_relColSizes[ci];
            }
        }
        d.ci("beginTableCell.colSpanSize:", colSpanSize);

        var fcurrentColXLSXIndex = currentColXLSXIndex;
        var fcolSpanSize = colSpanSize;
        IntStream.range(currentRowIndex, currentRowIndex + rowSpan).forEachOrdered(ri -> {
            IntStream.range(fcurrentColXLSXIndex, fcurrentColXLSXIndex + fcolSpanSize).forEachOrdered(ci -> {
                if (ri == currentRowIndex) {
                    table.setValue(ri, ci, CELL_Occupied);
                } else {
                    table.setValue(ri, ci, CELL_Occupied + " " + currentColRelIndex + " " + colSpan);
                }
                d.ci("beginTableCell.CELL_FULL ri:" + ri + ", ci:" + ci);
            });
        });
        lastCell = new TS_MIFXLSX_RichCell(xlsx, fileCommonConfig, rowSpan, colSpan, cellHeight, colSpanSize);
        table.setValue(currentRowIndex, currentColXLSXIndex, lastCell);
        return true;
    }

    private int calcColRelIndex(int rowIdx) {
        var emptyCellIdx = 0;
        if (isClosed()) {
            return emptyCellIdx;
        }
        String prevrelColIdx = null;
        for (var ci = 0; ci < table.getColumnSize(rowIdx); ci++) {
            d.ci("calcColRelIndex.ci:" + ci);
            var o = table.getValueAsObject(rowIdx, ci);
            switch (o) {
                case String ins -> {
                    if (CELL_Vacant.equals(ins)) {
                        d.ci("calcColRelIndex.iisString.isEmpty");
                        break;
                    } else {//starts with CELL_FULL
                        var tokens = TGS_StringUtils.jre().toList_spc(ins);
                        d.ci("calcColRelIndex.iisString.isFull.tokens:" + TGS_StringUtils.cmn().toString(tokens, ","));
                        if (tokens.size() == 1) {
                            d.ci("calcColRelIndex.iisString.isFull.skipped");
                        } else {
                            var relColIdx = tokens.get(1);
                            var relColSpan = tokens.get(2);
                            d.ci("calcColRelIndex.iisString.isFull.@colIdx:" + relColIdx + ".wColPan:" + relColSpan);
                            if (prevrelColIdx == null || !relColIdx.equals(prevrelColIdx)) {
                                prevrelColIdx = relColIdx;
                                emptyCellIdx += TGS_CastUtils.toInteger(relColSpan);
                                d.ci("calcColRelIndex.iisString.isFull.emptyCellIdx:" + emptyCellIdx);
                            }
                        }
                    }
                }
                case TS_MIFXLSX_RichCell ins -> {
                    emptyCellIdx += ins.colSpan;
                    d.ci("calcColRelIndex.iisString.isRich.emptyCellIdx:" + emptyCellIdx);
                }
                default ->
                    d.ci("calcColRelIndex.iisString.isERROR.o:" + o);
            }
        }
        return emptyCellIdx;
    }

    @Override
    public boolean endTableCell(int rotationInDegrees_0_90_180_270) {
        if (isClosed()) {
            return true;
        }
        lastCell.rotationInDegrees_0_90_180_270 = rotationInDegrees_0_90_180_270;
        var colSpanSize = 0;
        for (var ci = currentColRelIndex; ci < currentColRelIndex + lastCell.colSpan; ci++) {
            if (ci >= table_relColSizes.length) {
                colSpanSize++;
            } else {
                colSpanSize += table_relColSizes[ci];
            }
        }
        currentColRelIndex = calcColRelIndex(currentRowIndex);
//        currentColRelIndex += lastCell.colSpan;
        d.ci("endTableCell.currentColRelIndex:" + currentColRelIndex);
        d.ci("endTableCell.colSpanSize:" + colSpanSize);
        d.ci("endTableCell.currentColXLSXIndex:" + currentColXLSXIndex);
        lastCell = null;
        return true;
    }

    @Override
    public boolean beginText(int allign_Left0_center1_right2_just3) {
        if (isClosed()) {
            return true;
        }
        if (table == null) {
            if (lastCell != null) {
                d.ce("beginText why lastCell exists!");
                return false;
            }
            lastCell = new TS_MIFXLSX_RichCell(xlsx, fileCommonConfig, null, null, null, null);
            lastCell.setStyle(allign_Left0_center1_right2_just3, false);
            doc.add(lastCell);
            return true;
        } else {
            if (lastCell == null) {
                d.ce("beginText why lastCell not exists!");
                return false;
            }
            lastCell.setStyle(allign_Left0_center1_right2_just3, false);
            return true;
        }
    }

    @Override
    public boolean endText() {
        if (isClosed()) {
            return true;
        }
        if (table == null) {
            lastCell = null;
        }
        return true;
    }

    private TGS_UnionExcuseVoid saveFile_close() {
        return TGS_UnSafe.call(() -> {
            d.ci("compileFile.*** adding last line...");
            beginText(0);
            addText("");
            endText();

            d.ci("compileFile.*** xlsx.setPageSize(landscape, pageSizeAX);...");
            xlsx.setPageSize(landscape, pageSizeAX);
            var maxPageColumnCount = xlsx.getMaxPageColumnCount();

            d.ci("compileFile.*** double for.start...");

            currentRowIndex = 0;
            for (var di = 0; di < doc.size(); di++) {//INC
                d.ci("compileFile.*** for1.di:" + di);
                switch (doc.get(di)) {
                    case TS_MIFXLSX_RichCell cellDI -> {
                        var xlsxCell = xlsx.getCell(currentRowIndex, 0);
                        xlsx.setCellStyle(xlsxCell, xlsx.createCellStyle(cellDI.allign_center1_right2_defaultLeft, cellDI.isBordered));
                        xlsx.setCellRichText(xlsxCell, xlsx.createRichText(cellDI.texts, cellDI.fonts));
                        var ra = xlsx.createMergedCell(currentRowIndex, currentRowIndex, 0, maxPageColumnCount - 1, true);
                        if (cellDI.isBordered) {
                            xlsx.setBordersToMergedCell(ra, true);
                        }
                        var rowHeight = xlsx.calculateCellHeight(cellDI.texts, cellDI.fonts, maxPageColumnCount);
                        d.ci("compileFile. >>> MIFXLSX.parag.setRowHeight. ri:" + currentRowIndex + ", rh:" + rowHeight);
                        xlsx.setRowHeight(currentRowIndex, rowHeight);
                        currentRowIndex++;
                    }
                    case TGS_ListTable ins -> {
                        List<Integer> newHeightsWithRowSpan_value = TGS_ListUtils.of();
                        List<Integer> newHeightsWithRowSpan_live = TGS_ListUtils.of();
                        for (var ri = 0; ri < ins.getRowSize(); ri++) {
                            var maxRowHeight = 1;
                            for (var ci = 0; ci < ins.getColumnSize(ri); ci++) {//max
//                            d("compileFile.table.getValueAsObject(" + ri + ", " + ci + ") -> " + table.getValueAsObject(ri, ci));
                                var o = ins.getValueAsObject(ri, ci);
                                if (!(o instanceof TS_MIFXLSX_RichCell)) {
//                                d("compileFile.*** skipping cell :" + table.getValueAsObject(ri, ci));
                                    continue;
                                }
                                var cell = (TS_MIFXLSX_RichCell) o;
                                var xlsxCell = xlsx.getCell(currentRowIndex, ci);
                                xlsx.setCellStyle(xlsxCell, xlsx.createCellStyle(cell.allign_center1_right2_defaultLeft, true));
                                xlsx.setCellRichText(xlsxCell, xlsx.createRichText(cell.texts, cell.fonts));
                                var ra = xlsx.createMergedCell(currentRowIndex, currentRowIndex + cell.rowSpan - 1, ci, ci + cell.colSpanSize - 1, true);
                                if (cell.isBordered) {
                                    xlsx.setBordersToMergedCell(ra, true);
                                }
                                {
                                    var newHeight = xlsx.calculateCellHeight(cell.texts, cell.fonts, cell.colSpan) / cell.rowSpan;
                                    d.ci("compileFile. >>> MIFXLSX.table.calculateCellHeight.newHeight:" + newHeight);
                                    if (cell.rowSpan > 1) {
                                        d.ci("compileFile. >>> MIFXLSX.table.newHeightsWithRowSpan.add v:" + newHeight + ", rs:" + cell.rowSpan);
                                        newHeightsWithRowSpan_value.add(newHeight);
                                        newHeightsWithRowSpan_live.add(cell.rowSpan);
                                    } else {
                                        if (newHeight > maxRowHeight) {
                                            maxRowHeight = newHeight;
                                            d.ci("compileFile. >>> MIFXLSX.table.maxRowHeight is chaged:" + maxRowHeight);
                                        } else {
                                            d.ci("compileFile. >>> MIFXLSX.table.maxRowHeight is bigger:" + maxRowHeight);
                                        }
                                    }
                                }
                            }
                            for (var i = 0; i < newHeightsWithRowSpan_value.size(); i++) {//max
                                var value = newHeightsWithRowSpan_value.get(i);
                                var live = newHeightsWithRowSpan_live.get(i);
                                d.ci("compileFile. >>> MIFXLSX.table.newHeightsWithRowSpan[" + i + "] v:" + value + ", rs:" + live);
                                if (value > maxRowHeight) {
                                    maxRowHeight = value;
                                    d.ci("compileFile. >>> MIFXLSX.table.newHeightsWithRowSpan->maxRowHeight is chaged:" + maxRowHeight);
                                }
                                if (live == 1) {
                                    newHeightsWithRowSpan_value.remove(i);
                                    newHeightsWithRowSpan_live.remove(i);
                                    i--;
                                    d.ci("compileFile. >>> MIFXLSX.table.newHeightsWithRowSpan[" + i + "].removed->rs:" + live);
                                } else {
                                    newHeightsWithRowSpan_live.set(i, live - 1);
                                    d.ci("compileFile. >>> MIFXLSX.table.newHeightsWithRowSpan[" + i + "].decreasedBy1->rs:" + live);
                                }
                            }
                            d.ci("compileFile >>> MIFXLSX.table.setRowHeight. ri:" + currentRowIndex + ", rh:" + maxRowHeight);
                            xlsx.setRowHeight(currentRowIndex, maxRowHeight);
                            var fRi = ri;
                            IntStream.range(0, ins.getColumnSize(fRi)).forEachOrdered(ci -> {
                                var o = ins.getValueAsObject(fRi, ci);
                                if (!(o instanceof TS_MIFXLSX_RichCell)) {
//                                d("compileFile.*** skipping cell :" + table.getValueAsObject(ri, ci));
                                    return;
                                }
                                var cell = (TS_MIFXLSX_RichCell) o;
                                IntStream.range(0, cell.imgFiles.size()).forEachOrdered(i -> {
                                    d.ci("compileFile.*** addImage i:" + i);
                                    xlsx.addImage(cell.imgFiles.get(i).toAbsolutePath().toString(), currentRowIndex, ci, cell.colSpan);
                                });
                            });
                            currentRowIndex++;
                        }
                    }
                    default ->
                        d.ci("compileFile *** MIFXLSX.ERROR. unknown doc.get object: " + doc.get(di));
                }
            }

            d.ci("compileFile.*** double xlsx.close();...");
            xlsx.close();
            return TGS_UnionExcuseVoid.ofVoid();
        }, e -> {
            TGS_UnSafe.run(() -> xlsx.close(), e2 -> {
                //DO NOTHING
            });
            return TGS_UnionExcuseVoid.ofExcuse(e);
        });
    }

    private static class TS_MIFXLSX_RichCell {

        List<String> texts;
        List<Font> fonts;
        TS_FileXlsxUtils xlsx;
        TS_FileCommonConfig fileCommonConfig;
        int rowIdx, colIdx;
        Integer rowSpan, colSpan, rowHeight;
        Integer colSpanSize;
        boolean isTableCell;
        boolean isBordered;
        int allign_center1_right2_defaultLeft;
        int rotationInDegrees_0_90_180_270 = 0;
        List<Path> imgFiles;

        public TS_MIFXLSX_RichCell(TS_FileXlsxUtils xlsx, TS_FileCommonConfig fileCommonConfig, Integer rowSpan, Integer colSpan, Integer rowHeight, Integer colSpanSize) {
            this.fileCommonConfig = fileCommonConfig;
            this.xlsx = xlsx;
            this.isTableCell = rowSpan != null;
            this.rowSpan = rowSpan;
            this.colSpan = colSpan;
            this.colSpanSize = colSpanSize;
            this.rowHeight = rowHeight;
            texts = TGS_ListUtils.of();
            fonts = TGS_ListUtils.of();
            imgFiles = TGS_ListUtils.of();

        }

        public void setStyle(int allign_center1_right2_defaultLeft, boolean isBordered) {
            this.allign_center1_right2_defaultLeft = allign_center1_right2_defaultLeft;
            this.isBordered = isTableCell || isBordered;
        }

    }
}
