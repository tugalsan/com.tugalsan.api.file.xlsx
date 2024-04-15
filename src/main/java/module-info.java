module com.tugalsan.api.file.xlsx {
    requires java.desktop;
    requires poi;
    requires poi.ooxml;
    requires com.tugalsan.api.list;
    requires com.tugalsan.api.url;
    requires com.tugalsan.api.charset;
    requires com.tugalsan.api.cast;
    requires com.tugalsan.api.union;
    requires com.tugalsan.api.validator;
    requires com.tugalsan.api.math;
    requires com.tugalsan.api.coronator;
    requires com.tugalsan.api.callable;
    requires com.tugalsan.api.runnable;
    requires com.tugalsan.api.log;
    requires com.tugalsan.api.string;
    requires com.tugalsan.api.file;
    requires com.tugalsan.api.file.html;
    requires com.tugalsan.api.file.common;
    exports com.tugalsan.api.file.xlsx.server;
}
