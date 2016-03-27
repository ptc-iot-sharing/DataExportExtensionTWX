package com.ptc.thingworx;

import com.lowagie.text.Document;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import com.thingworx.entities.utils.ThingUtilities;
import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.InfoTable;
import com.thingworx.types.primitives.IPrimitiveType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.ByteArrayOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;


class ExporterUtils {

    //setting the datetime format that suits our need
    private DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
    private Calendar cal = Calendar.getInstance();

    private FileRepositoryThing DataExporterRepository;

    ExporterUtils() {

        //search for the repository thing and call one of its method to create the directory inside the Thingworx Storage
        DataExporterRepository = (FileRepositoryThing) ThingUtilities.findThing("DataExporterRepository");
    }

    String ExportInfotableAsPdf(InfoTable infotable) throws Exception {

        //identify how many columns should the table have
        int columnSize = infotable.getRow(0).size();

        Document document = new Document();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        String fileName = "Export" + dateFormat.format(cal.getTime()) + ".docx";

        PdfWriter.getInstance(document, bos);
        document.open();
        //create the pdf file with the right numbers of columns
        PdfPTable table = new PdfPTable(columnSize);
        table.setSplitLate(false);

        //create the table header with the field definitions from the input infotable
        infotable.getRow(0).keySet().forEach(table::addCell);

        //populate the rest of the table with values
        for (int i = 0; i < infotable.getRowCount(); i++)
            for (IPrimitiveType value : infotable.getRow(i).values()) table.addCell(value.getStringValue());
        document.add(table);
        DataExporterRepository.CreateBinaryFile(fileName, bos.toByteArray(), true);

        document.close();
        return DataExporterRepository.GetFileListingWithLinks("/", fileName).getRow(0).getStringValue("downloadLink");
    }

    String ExportInfotableAsExcel(InfoTable infotable) throws Exception {

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        new InfotableToExcelPoi().generate(infotable, bos);

        String fileName = "Export" + dateFormat.format(cal.getTime()) + ".xlsx";
        DataExporterRepository.CreateBinaryFile(fileName, bos.toByteArray(), true);
        bos.close();

        return DataExporterRepository.GetFileListingWithLinks("/", fileName).getRow(0).getStringValue("downloadLink");
    }


    String ExportInfotableAsWord(InfoTable infotable) throws Exception {
        XWPFDocument document = new XWPFDocument();
        XWPFTable table = document.createTable();

        //create the table header with the field definitions from the input infotable
        XWPFTableRow tableRowOne = table.getRow(0);
        int j = 0;
        for (String name : infotable.getRow(0).keySet()) {
            if (j > 0) {
                tableRowOne.addNewTableCell().setText(name);
            } else {
                tableRowOne.getCell(0).setText(name);
            }
            j++;
        }

        //populate the rest of the table with values
        for (int i = 0; i < infotable.getRowCount(); i++) {
            XWPFTableRow tableNextRow = table.createRow();
            int x = 0;
            for (IPrimitiveType value : infotable.getRow(i).values()) {
                tableNextRow.getCell(x).setText(value.getStringValue());
                x++;
            }

        }
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        document.write(bos);
        document.close();

        String fileName = "Export" + dateFormat.format(cal.getTime()) + ".docx";
        DataExporterRepository.CreateBinaryFile(fileName, bos.toByteArray(), true);
        bos.close();

        return DataExporterRepository.GetFileListingWithLinks("/", fileName).getRow(0).getStringValue("downloadLink");
    }
}
