package com.ptc.thingworx;


import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.resources.Resource;
import com.thingworx.types.InfoTable;


public class Exporter extends Resource {

    protected static final ch.qos.logback.classic.Logger _Logger = LogUtilities.getInstance()
            .getApplicationLogger(Exporter.class);

    @ThingworxServiceDefinition(name = "ExportInfotableAsPdf", description = "")
    @ThingworxServiceResult(
            name = "result",
            description = "Link to the file",
            baseType = "STRING"
    )
    public String ExportInfotableAsPdf(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE") InfoTable infotable)
            throws Exception {
        ExporterUtils exp = new ExporterUtils();
        return exp.ExportInfotableAsPdf(infotable);

    }

    @ThingworxServiceDefinition(name = "ExportInfotableAsExcel", description = "")
    @ThingworxServiceResult(
            name = "result",
            description = "Link to the file",
            baseType = "STRING"
    )
    public String ExportInfotableAsExcel(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE") InfoTable infotable)
            throws Exception {
        ExporterUtils exp = new ExporterUtils();
        return exp.ExportInfotableAsExcel(infotable);

    }


    @ThingworxServiceDefinition(name = "ExportInfotableAsWord", description = "")
    @ThingworxServiceResult(
            name = "result",
            description = "Link to the file",
            baseType = "STRING"
    )
    public String ExportInfotableAsWord(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE") InfoTable infotable)
            throws Exception {
        ExporterUtils exp = new ExporterUtils();
        return exp.ExportInfotableAsWord(infotable);

    }

}