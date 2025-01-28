module com.tugalsan.api.file.xlsx {
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires org.apache.poi.ooxml.schemas;
    requires org.apache.poi.scratchpad;

    requires java.desktop;
    requires com.tugalsan.api.list;
    requires com.tugalsan.api.url;
    requires com.tugalsan.api.charset;
    requires com.tugalsan.api.union;
    requires com.tugalsan.api.cast;
    requires com.tugalsan.api.unsafe;
    requires com.tugalsan.api.math;
    requires com.tugalsan.api.function;
    requires com.tugalsan.api.log;
    requires com.tugalsan.api.string;
    requires com.tugalsan.api.file;
    requires com.tugalsan.api.file.html;
    requires com.tugalsan.api.file.common;
    exports com.tugalsan.api.file.xlsx.server;
}
