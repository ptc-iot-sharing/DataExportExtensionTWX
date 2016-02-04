package com.ptc.thingworx;


import org.slf4j.Logger;

import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.annotations.ThingworxConfigurationTableDefinitions;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.things.Thing;
import com.thingworx.types.InfoTable;

@SuppressWarnings("serial")

public class Exporter extends Thing {

    protected static ch.qos.logback.classic.Logger _Logger = LogUtilities.getInstance()
            .getApplicationLogger(Exporter.class);




    protected void initializeThing() throws Exception {



    }



    @ThingworxServiceDefinition(name = "ExportInfotableAsPdf", description = "")
    public void ExportInfotableAsPdf(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE") InfoTable infotable)
            throws Exception {
        ExporterUtils exp = new ExporterUtils();
        exp.ExportInfotableAsPdf(infotable);

    }
    @ThingworxServiceDefinition(name = "ExportInfotableAsExcel", description = "")
    public void ExportInfotableAsExcel(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE") InfoTable infotable)
            throws Exception {
        ExporterUtils exp = new ExporterUtils();
        exp.ExportInfotableAsExcel(infotable);

    }


    @ThingworxServiceDefinition(name = "ExportInfotableAsWord", description = "")
    public void ExportInfotableAsWord(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE") InfoTable infotable)
            throws Exception {
        ExporterUtils exp = new ExporterUtils();
        exp.ExportInfotableAsWord(infotable);

    }

}