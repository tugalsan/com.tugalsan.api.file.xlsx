package com.tugalsan.api.file.xlsx.server;

import com.tugalsan.api.charset.client.TGS_CharSetCast;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.IntStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.tugalsan.api.list.client.TGS_ListTable;
import com.tugalsan.api.file.html.client.TGS_FileHtmlUtils;
import com.tugalsan.api.unsafe.client.*;

public class TS_FileXlsxTable extends TGS_ListTable {
    
    private TS_FileXlsxTable(){
        super(true);
    }
    
    public static TS_FileXlsxTable of(){
        return new TS_FileXlsxTable();
    }

    public boolean toFile(Path destXLSX) {
        return toFile(this, destXLSX);
    }

    public static TS_FileXlsxTable ofXlsx() {
        return new TS_FileXlsxTable();
    }

    public static boolean toFile(TGS_ListTable table, Path destXLSX) {
        try (var xlsx = new TS_FileXlsx(destXLSX);) {
            var fBold = xlsx.createFont(true, false, false);
            var fPlain = xlsx.createFont(false, false, false);
            IntStream.range(0, table.getRowSize()).forEachOrdered(ri -> {
                IntStream.range(0, table.getColumnSize(ri)).forEachOrdered(ci -> {
                    xlsx.setCellRichText(xlsx.getCell(ri, ci), xlsx.createRichText(table.getValueAsString(ri, ci), ri == 0 ? fBold : fPlain));
                });
            });
        }
        return Files.exists(destXLSX) && !Files.isDirectory(destXLSX);
    }

    public static StringBuffer toHTML(Path destXLSX, CharSequence optionalCustomDomain) {
        return TGS_UnSafe.call(() -> {
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
                        sb.append(TGS_FileHtmlUtils.beginLines(destXLSX.toString(), true, false, 5, 5, null, false, optionalCustomDomain));
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
                        return sb;
                    }
                }
                try (var workbook = new XSSFWorkbook(is);) {
                    var sb = new StringBuffer();
                    sb.append(TGS_FileHtmlUtils.beginLines(destXLSX.toString(), true, false, 5, 5, null, false, optionalCustomDomain));
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
                    return sb;
                }
            }
        });
    }
}
