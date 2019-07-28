package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.*;

abstract class BasisBuilder {

    protected final Workbook wb;
    protected final InputParams params;
    protected final Sheet sheet;
    protected final CellStyle headerStyle;
    protected final CellStyle dataStyle;
    protected final CellStyle timeStyle;


    public BasisBuilder(Workbook wb, InputParams params, int indexSheet) {

        this.wb = wb;
        this.params = params;
        this.sheet = wb.createSheet();
        wb.setSheetName(indexSheet, getSheetName());
        headerStyle = generateHeaderStyle();
        dataStyle = generateDataStyle();
        timeStyle = generateTimeFormat();
    }

    abstract String getSheetName();

    public abstract void build();

    private CellStyle generateHeaderStyle() {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderTop(BorderStyle.MEDIUM);
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        return cellStyle;
    }

    private CellStyle generateDataStyle() {
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    private CellStyle generateTimeFormat() {
        final CellStyle cellStyle = generateDataStyle();
        final DataFormat dataFormat = wb.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("H:MM:SS;@"));
        return cellStyle;
    }
}
