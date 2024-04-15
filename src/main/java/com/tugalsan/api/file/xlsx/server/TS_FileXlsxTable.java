package com.tugalsan.api.file.xlsx.server;

import com.tugalsan.api.charset.client.TGS_CharSetCast;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.tugalsan.api.list.client.TGS_ListTable;
import com.tugalsan.api.file.html.client.TGS_FileHtmlUtils;
import com.tugalsan.api.union.client.TGS_UnionExcuse;
import com.tugalsan.api.union.client.TGS_UnionExcuseVoid;
import com.tugalsan.api.url.client.TGS_Url;
import java.io.IOException;

public class TS_FileXlsxTable extends TGS_ListTable {

    private TS_FileXlsxTable() {
        super(true);
    }

    public TGS_UnionExcuseVoid toFile(Path destXLSX) {
        return toFile(this, destXLSX);
    }

    public static TS_FileXlsxTable ofXlsx() {
        return new TS_FileXlsxTable();
    }

    public static TGS_UnionExcuseVoid toFile(TGS_ListTable table, Path destXLSX) {
        try (var xlsx = new TS_FileXlsxUtils(destXLSX);) {
            var fBold = xlsx.createFont(true, false, false);
            var fPlain = xlsx.createFont(false, false, false);
            for (var ri = 0; ri < table.getRowSize(); ri++) {
                for (var ci = 0; ci < table.getColumnSize(ri); ci++) {
                    var u = table.getValueAsString(ri, ci);
                    if (u.isExcuse()) {
                        return u.toExcuseVoid();
                    }
                    xlsx.setCellRichText(
                            xlsx.getCell(ri, ci),
                            xlsx.createRichText(u.value(), ri == 0 ? fBold : fPlain)
                    );
                }
            }
        }
        if (Files.exists(destXLSX) && !Files.isDirectory(destXLSX)) {
            return TGS_UnionExcuseVoid.ofVoid();
        } else {
            return TGS_UnionExcuseVoid.ofExcuse(TS_FileXlsxTable.class.getSimpleName(), "toFile", "NOT: Files.exists(destXLSX) && !Files.isDirectory(destXLSX)");
        }
    }

    public static TGS_UnionExcuse<StringBuffer> toHTML(Path destXLSX, TGS_Url bootLoaderJs) {
        try (var is = Files.newInputStream(destXLSX)) {
            var FILE_TYPES = new String[]{"xls", "xlsx"};
            var NEW_LINE = "\n";
            var HTML_TR_S = "<tr>";
            var HTML_TR_E = "</tr>";
            var HTML_TD_S = "<td>";
            var HTML_TD_E = "</td>";
            var isXLS = TGS_CharSetCast.toLocaleLowerCase(destXLSX.toAbsolutePath().toString()).endsWith(FILE_TYPES[0]);
            if (isXLS) {
                try (var workbook = new HSSFWorkbook(is);) {
                    var sb = new StringBuffer();
                    sb.append(TGS_FileHtmlUtils.beginLines(destXLSX.toString(), false, 5, 5, null, false, bootLoaderJs));
                    sb.append("<table>");
                    var sheet = workbook.getSheetAt(0);
                    var rows = sheet.rowIterator();
                    rows.forEachRemaining(row -> {
                        var cells = row.cellIterator();
                        sb.append(NEW_LINE);
                        sb.append(HTML_TR_S);
                        cells.forEachRemaining(cell -> {
                            sb.append(HTML_TD_S);
                            sb.append(cell.toString());
                            sb.append(HTML_TD_E);
                        });
                        sb.append(HTML_TR_E);
                    });
                    sb.append(NEW_LINE);
                    sb.append("<table>");
                    sb.append(TGS_FileHtmlUtils.endLines(false));
                    return TGS_UnionExcuse.of(sb);
                }
            }
            try (var workbook = new XSSFWorkbook(is);) {
                var sb = new StringBuffer();
                sb.append(TGS_FileHtmlUtils.beginLines(destXLSX.toString(), false, 5, 5, null, false, bootLoaderJs));
                sb.append("<table>");
                var sheet = workbook.getSheetAt(0);
                var rows = sheet.rowIterator();
                rows.forEachRemaining(row -> {
                    var cells = row.cellIterator();
                    sb.append(NEW_LINE);
                    sb.append(HTML_TR_S);
                    cells.forEachRemaining(cell -> {
                        sb.append(HTML_TD_S);
                        sb.append(cell.toString());
                        sb.append(HTML_TD_E);
                    });
                    sb.append(HTML_TR_E);
                });
                sb.append(NEW_LINE);
                sb.append("<table>");
                sb.append(TGS_FileHtmlUtils.endLines(false));
                return TGS_UnionExcuse.of(sb);
            }
        } catch (IOException ex) {
            return TGS_UnionExcuse.ofExcuse(ex);
        }
    }
}
