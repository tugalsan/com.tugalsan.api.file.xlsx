module com.tugalsan.api.file.xlsx {
    requires java.desktop;
    requires poi;
    requires poi.ooxml;
    requires ooxml.schemas;
    requires com.tugalsan.api.list;
    requires com.tugalsan.api.executable;
    requires com.tugalsan.api.log;
    requires com.tugalsan.api.string;
    requires com.tugalsan.api.file;
    requires com.tugalsan.api.file.html;
    exports com.tugalsan.api.file.xlsx.server;
}