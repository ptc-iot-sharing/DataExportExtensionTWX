package com.ptc.thingworx;

import com.lowagie.text.Font;
import com.lowagie.text.Rectangle;
import com.lowagie.text.*;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import com.thingworx.entities.utils.ThingUtilities;
import com.thingworx.metadata.FieldDefinition;
import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.InfoTable;
import com.thingworx.types.primitives.IPrimitiveType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;

import java.awt.*;
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
        int columnSize = infotable.getFieldCount();

        final Rectangle pageSize = PageSize.A4.rotate();
        Document document = new Document(pageSize);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        String fileName = "Export" + dateFormat.format(cal.getTime()) + ".pdf";

        final PdfWriter writer = PdfWriter.getInstance(document, bos);
        document.open();
        //create the pdf file with the right numbers of columns
        PdfPTable table = new PdfPTable(columnSize);
        table.setTotalWidth(pageSize.getWidth() - 10);
        table.setLockedWidth(true);
        table.setSplitLate(false);
        Font headerFont = new Font(Font.TIMES_ROMAN, 11, Font.BOLD, new Color(0, 0, 0));
        Font rowFont = new Font(Font.TIMES_ROMAN, 11);

        //create the table header with the field definitions from the input infotable
        infotable.getDataShape().getFields().getOrderedFieldsByOrdinal().forEach(
                fieldDefinition -> {
                    PdfPCell cell = new PdfPCell(new Phrase(fieldDefinition.getName(), headerFont));
                    cell.setBackgroundColor(Color.LIGHT_GRAY);
                    table.addCell(cell);
                });

        //populate the rest of the table with values
        for (int rowIndex = 0; rowIndex < infotable.getRowCount(); rowIndex++) {
            table.completeRow();
            for (FieldDefinition field : infotable.getDataShape().getFields().getOrderedFieldsByOrdinal()) {
                IPrimitiveType cellValue = infotable.getRow(rowIndex).getOrDefault(field.getName(),
                        field.getDefaultValue());
                if (cellValue != null) {
                    PdfPCell cell = getPdfCell(cellValue, FormatType.getFormatType(field.getBaseType()), rowFont);
                    table.addCell(cell);
                } else {
                    table.addCell("");

                }
            }
        }
        document.add(table);

        document.close();
        DataExporterRepository.CreateBinaryFile(fileName, bos.toByteArray(), true);

        return DataExporterRepository.GetFileListingWithLinks("/", fileName).getRow(0).getStringValue("downloadLink");
    }

    private PdfPCell getPdfCell(IPrimitiveType value, FormatType formatType, Font font) {
        PdfPCell cell = new PdfPCell();
        if (value != null) {
            switch (formatType) {
                case HEADER:
                    cell.setPhrase(new Phrase(value.getStringValue(), font));
                    cell.setBackgroundColor(Color.LIGHT_GRAY);
                    break;
                case TEXT:
                    cell.setPhrase(new Phrase(value.getStringValue(), font));
                    break;
                case INTEGER:
                    cell.setPhrase(new Phrase(String.format("%d", (Integer) value.getValue()), font));
                    break;
                case FLOAT:
                    cell.setPhrase(new Phrase(String.format("%.00f", (Double) value.getValue()), font));
                    break;
                case DATE:
                    DateTimeFormatter dtf = DateTimeFormat.forPattern("MM/dd/yyyy HH:mm:ss");
                    cell.setPhrase(new Phrase(dtf.print((DateTime) value.getValue()), font));
                    break;
            }
        }
        return cell;
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
        for (FieldDefinition fieldDefinition : infotable.getDataShape().getFields().getOrderedFieldsByOrdinal()) {
            if (j > 0) {
                tableRowOne.addNewTableCell().setText(fieldDefinition.getName());
            } else {
                tableRowOne.getCell(0).setText(fieldDefinition.getName());
            }
            j++;
        }

        //populate the rest of the table with values
        for (int rowIndex = 0; rowIndex < infotable.getRowCount(); rowIndex++) {
            XWPFTableRow tableNextRow = table.createRow();
            int x = 0;
            for (FieldDefinition field : infotable.getDataShape().getFields().getOrderedFieldsByOrdinal()) {
                IPrimitiveType cellValue = infotable.getRow(rowIndex).getOrDefault(field.getName(),
                        field.getDefaultValue());
                if (cellValue != null) {
                    tableNextRow.getCell(x).setText(cellValue.getStringValue());
                } else {
                    tableNextRow.getCell(x).setText("");
                }
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
