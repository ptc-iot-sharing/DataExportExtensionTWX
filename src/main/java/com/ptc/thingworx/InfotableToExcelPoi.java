package com.ptc.thingworx;

import com.thingworx.common.utils.StringUtilities;
import com.thingworx.types.BaseTypes;
import com.thingworx.types.InfoTable;
import com.thingworx.types.primitives.IPrimitiveType;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.text.ParsePosition;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * IntelliPBX - (c) Lacatus Petrisor, 2014
 * Writes a TableModel to an XLS file.
 */
public class InfotableToExcelPoi {
    private final static Logger LOGGER = LoggerFactory.getLogger(InfotableToExcelPoi.class.getName());

    private Workbook workbook;
    private Sheet sheet;
    private Font boldFont;
    private DataFormat format;
    private InfoTable infoTable;
    private FormatType[] formatTypes;
    private CellStyle timeIntervalCellStyle;
    private SimpleDateFormat timeIntervalFormat;
    private CellStyle linkCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle percentStyle;
    private SimpleDateFormat dateFormat;


    private void createTableModelToExcel(Workbook workbook, InfoTable tableModel) {
        this.workbook = workbook;
        this.infoTable = tableModel;
        sheet = workbook.createSheet();
        createCellStyles(workbook);
    }

    private void createCellStyles(Workbook workbook) {
        boldFont = workbook.createFont();
        boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        format = workbook.createDataFormat();
        timeIntervalCellStyle = workbook.createCellStyle();
        timeIntervalCellStyle.setDataFormat(workbook.getCreationHelper().
                createDataFormat().getFormat("[HH]:MM:SS.00"));
        timeIntervalFormat = new SimpleDateFormat("HH:mm:ss.SSS");
        linkCellStyle = workbook.createCellStyle();
        // we have a hyperlink
        Font linkFont = workbook.createFont();
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkCellStyle.setFont(linkFont);
        dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(
                workbook.getCreationHelper().createDataFormat().getFormat("DD-MM-YYYY HH:MM:SS.000"));
        percentStyle = workbook.createCellStyle();
        percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.0%"));
        dateFormat = new SimpleDateFormat("dd-MM-yyyy HH:mm:ss.SSS");
    }

    private FormatType getFormatType(BaseTypes baseType) {
        if (baseType == BaseTypes.INTEGER || baseType == BaseTypes.LONG) {
            return FormatType.INTEGER;
        } else if (baseType == BaseTypes.NUMBER) {
            return FormatType.FLOAT;
        } else if (baseType == BaseTypes.DATETIME) {
            return FormatType.DATE;
        } else {
            return FormatType.TEXT;
        }
    }

    public void generate(InfoTable infoTable, OutputStream outputStream) throws IOException {
        try {
            createTableModelToExcel(new XSSFWorkbook(), infoTable);
            this.infoTable = infoTable;
            generate();
            workbook.write(outputStream);
        } finally {
            outputStream.close();
        }
    }

    public void generate() {
        LOGGER.info("Started report generation");
        if (formatTypes != null && formatTypes.length != infoTable.getFieldCount()) {
            throw new IllegalStateException("Number of types is not identical to number of infoTable columns. " +
                    "Number of types: " + formatTypes.length + ". Number of columns: " + infoTable.getFieldCount());
        }

        int currentRow = 0;
        Row row = sheet.createRow(currentRow);
        int numCols = infoTable.getFieldCount();
        int numRows = infoTable.getRowCount();
        boolean isAutoDecideFormatTypes;
        if (isAutoDecideFormatTypes = (formatTypes == null)) {
            formatTypes = new FormatType[numCols];
        }
        int i = 0;
        for (String name : infoTable.getFirstRow().keySet()) {
            writeCell(row, i, name, FormatType.HEADER, boldFont);
            if (isAutoDecideFormatTypes) {
                formatTypes[i] = getFormatType(infoTable.getField(name).getBaseType());
            }
            i++;

        }
        autoSizeColumns(numCols);
        LOGGER.info("Written table header");
        currentRow++;

        // Write report rows
        for (int rowIndex = 0; rowIndex < numRows; rowIndex++) {
            try {
                row = sheet.createRow(currentRow++);
            } catch (IllegalArgumentException ex) {
                LOGGER.info("Attempted to create a file with more than the maximum" +
                        " allowed number of rows. Overflow remaining rows to a new sheet.");
                sheet = workbook.createSheet();
                currentRow = 0;
                row = sheet.createRow(currentRow++);
            }
            try {
                int colIndex = 0;
                for (IPrimitiveType value : infoTable.getRow(rowIndex).values()) {
                    writeCell(row, colIndex, value.getStringValue(), formatTypes[colIndex], null);
                    colIndex++;
                }
            } catch (Exception ex) {
                LOGGER.error("Failed to write the item at " + row.getRowNum(), ex);
            }
            // after some data has been written, auto-size again
            if (rowIndex <= 5) {
                // Auto-size columns
                autoSizeColumns(numCols);
            }
        }
        LOGGER.info("Written all cells");
    }

    /**
     * Draws border around a given cell
     * // WARNING: for large files this is extremely slow
     *
     * @param cell cell to draw borders around
     */
    private void drawBorders(Cell cell) {
        CellUtil.setCellStyleProperty(cell, workbook, CellUtil.BORDER_BOTTOM, CellStyle.BORDER_THIN);
        CellUtil.setCellStyleProperty(cell, workbook, CellUtil.BORDER_LEFT, CellStyle.BORDER_THIN);
        CellUtil.setCellStyleProperty(cell, workbook, CellUtil.BORDER_RIGHT, CellStyle.BORDER_THIN);
        CellUtil.setCellStyleProperty(cell, workbook, CellUtil.BORDER_TOP, CellStyle.BORDER_THIN);
    }

    private void autoSizeColumns(int numCols) {
        // Auto-size columns
        for (int i = 0; i < numCols; i++) {
            sheet.autoSizeColumn((short) i);
        }
    }

    private void writeCell(Row row, int col, Object value, FormatType formatType, Font font) {
        writeCell(row, col, value, formatType, null, font);
    }

    private void writeCell(Row row, int col, Object value, FormatType formatType,
                           Short bgColor, Font font) {
        Cell cell = CellUtil.createCell(row, col, null);
        if (value != null) {
            if (font != null) {
                CellStyle style = workbook.createCellStyle();
                style.setFont(font);
                cell.setCellStyle(style);
            }
            switch (formatType) {
                case HEADER:
                    cell.setCellValue(value.toString());
                    CellUtil.setCellStyleProperty(cell, workbook,
                            CellUtil.FILL_FOREGROUND_COLOR, HSSFColor.ORANGE.index);
                    CellUtil.setCellStyleProperty(cell, workbook,
                            CellUtil.FILL_PATTERN, CellStyle.SOLID_FOREGROUND);
                    break;
                case TEXT:
                    writeTextCell(value, cell);
                    break;
                case INTEGER:
                    cell.setCellValue(((Number) value).intValue());
                    CellUtil.setCellStyleProperty(cell, workbook, CellUtil.DATA_FORMAT,
                            HSSFDataFormat.getBuiltinFormat(("#,##0")));
                    break;
                case FLOAT:
                    cell.setCellValue(((Number) value).doubleValue());
                    CellUtil.setCellStyleProperty(cell, workbook, CellUtil.DATA_FORMAT,
                            HSSFDataFormat.getBuiltinFormat(("#,##0.00")));
                    break;
                case DATE:
                    cell.setCellValue((Timestamp) value);
                    CellUtil.setCellStyleProperty(cell, workbook, CellUtil.DATA_FORMAT,
                            HSSFDataFormat.getBuiltinFormat(("m/d/yy")));
                    break;
                case MONEY:
                    cell.setCellValue(((Number) value).intValue());
                    CellUtil.setCellStyleProperty(cell, workbook,
                            CellUtil.DATA_FORMAT, format.getFormat("($#,##0.00);($#,##0.00)"));
                    break;
                case PERCENTAGE:
                    cell.setCellValue(((Number) value).doubleValue());
                    CellUtil.setCellStyleProperty(cell, workbook,
                            CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("0.00%"));
            }
        }

        if (row.getRowNum() < 1_000) {
            drawBorders(cell);
        }
        if (bgColor != null) {
            CellUtil.setCellStyleProperty(cell, workbook, CellUtil.FILL_FOREGROUND_COLOR, bgColor);
            CellUtil.setCellStyleProperty(cell, workbook, CellUtil.FILL_PATTERN, CellStyle.SOLID_FOREGROUND);
        }
    }

    private void writeTextCell(Object value, Cell cell) {
        String stringValue = value.toString();
        boolean isAlertString = false;
        Date dateRepresentation;

        if (stringValue.startsWith("!!")) {
            stringValue = stringValue.substring(2);
            isAlertString = true;
        }

        if (stringValue.endsWith("%")) {
            cell.setCellStyle(percentStyle);
            stringValue = stringValue.substring(0, stringValue.length() - 1);
            if (StringUtilities.isNumber(stringValue)) {
                cell.setCellValue(new Double(stringValue) / 100);
            } else {
                cell.setCellValue(stringValue);
            }
        } else {
            if (StringUtilities.isNumber(stringValue)) {
                cell.setCellValue(new Double(stringValue));

                // we use DateFormat.parse(String,ParsePosition) in order to skip generating exceptions
            } else if ((dateRepresentation = timeIntervalFormat.
                    parse(stringValue, new ParsePosition(0))) != null) {
                cell.setCellStyle(timeIntervalCellStyle);
                Calendar cal = Calendar.getInstance();
                cal.setTime(dateRepresentation);
                cal.set(Calendar.YEAR, 1900);
                // we subtract 1.0 from the excel date to get a 00: based date.
                // Leaving it as it is would give us an offset of 24 hours
                cell.setCellValue(DateUtil.getExcelDate(cal.getTime()) - 1.0);

            } else if (stringValue.startsWith("http://")) {
                // we have a hyperlink
                cell.setCellStyle(linkCellStyle);
                cell.setCellValue("Ctrl+Click to open");
                Hyperlink link = workbook.getCreationHelper().createHyperlink(Hyperlink.LINK_URL);
                link.setAddress(stringValue);
                cell.setHyperlink(link);

            } else if ((dateRepresentation = dateFormat.
                    parse(stringValue, new ParsePosition(0))) != null) {
                // the regex match is not intended to be precise, just get a matched format
                cell.setCellStyle(dateCellStyle);
                cell.setCellValue(dateRepresentation);
            } else {
                cell.setCellValue(stringValue);
            }
        }
        if (isAlertString) {
            CellUtil.setCellStyleProperty(cell, workbook,
                    CellUtil.FILL_FOREGROUND_COLOR, HSSFColor.ORANGE.index);
            CellUtil.setCellStyleProperty(cell, workbook,
                    CellUtil.FILL_PATTERN, CellStyle.SOLID_FOREGROUND);
        }
    }

    private enum FormatType {
        TEXT,
        INTEGER,
        FLOAT,
        DATE,
        MONEY,
        PERCENTAGE,
        HEADER
    }
}