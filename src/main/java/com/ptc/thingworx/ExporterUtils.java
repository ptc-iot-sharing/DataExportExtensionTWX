package com.ptc.thingworx;

import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;


import com.thingworx.entities.utils.ThingUtilities;


import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.primitives.IPrimitiveType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import com.thingworx.types.InfoTable;


public class ExporterUtils {

    //setting the datetime format that suits our need
    DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
    Calendar cal = Calendar.getInstance();

    FileRepositoryThing DataExporterRepository;

    public ExporterUtils() {

        //search for the repository thing and call one of its method to create the directory inside the Thingworx Storage
        DataExporterRepository = (FileRepositoryThing) ThingUtilities.findThing("DataExporterRepository");
        try {
            DataExporterRepository.GetDirectoryStructure();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    public void ExportInfotableAsPdf(InfoTable infotable) throws DocumentException, ClassNotFoundException, IOException {

        //identify how many columns should the table have
        int columnSize = infotable.getRow(0).size();

        Document document = new Document();
        File file = new File(DataExporterRepository.getRootPath() + File.separator + "Export" + dateFormat.format(cal.getTime()) + ".pdf");

        PdfWriter.getInstance(document, new FileOutputStream(file));
        document.open();
        //create the pdf file with the right numbers of columns
        PdfPTable table = new PdfPTable(columnSize);
        table.setSplitLate(false);

        //create the table header with the field definitions from the input infotable
        for (String name : infotable.getRow(0).keySet()) table.addCell(name);

        //populate the rest of the table with values
        for (int i = 0; i < infotable.getRowCount(); i++)
            for (IPrimitiveType value : infotable.getRow(i).values()) table.addCell(value.getStringValue());
        document.add(table);

        document.close();
    }

    public void ExportInfotableAsExcel(InfoTable infotable) throws DocumentException, ClassNotFoundException, IOException {
        Workbook workbook = new XSSFWorkbook();
        //create the sheet with data
        Sheet sheet = workbook.createSheet("data");

        //create the table header with the field definitions from the input infotable
        Row rowhead = sheet.createRow((short) 0);
        int i = 0;
        for (String name : infotable.getRow(0).keySet()) {
            rowhead.createCell((short) i).setCellValue(name);
            sheet.autoSizeColumn(i);
            i++;
        }

        //populate the rest of the table with values
        for (int j = 0; j < infotable.getRowCount(); j++) {
            Row row = sheet.createRow((short) j + 1);
            int x = 0;
            for (IPrimitiveType value : infotable.getRow(j).values()) {
                row.createCell((short) x).setCellValue(value.getStringValue());
                x++;
                sheet.autoSizeColumn(x);
            }
        }

        File excelFile = new File(DataExporterRepository.getRootPath() + File.separator + "Export" + dateFormat.format(cal.getTime()) + ".xlsx");
        FileOutputStream fileOut = new FileOutputStream(excelFile);

        workbook.write(fileOut);
        workbook.close();
        fileOut.close();

    }


    public void ExportInfotableAsWord(InfoTable infotable) throws DocumentException, ClassNotFoundException, IOException {
        XWPFDocument document = new XWPFDocument();
        FileOutputStream out;
        out = new FileOutputStream(
                new File(DataExporterRepository.getRootPath() + File.separator + "Export" + dateFormat.format(cal.getTime()) + ".docx"));
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
        document.write(out);
        out.close();
    }
}
