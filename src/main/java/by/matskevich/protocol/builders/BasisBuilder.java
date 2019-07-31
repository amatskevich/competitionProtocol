package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.model.PlacesModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

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

    public abstract PlacesModel build();

    protected PlacesModel generateData(int rowIndex) {

        PlacesModel addresses = new PlacesModel(getSheetName());
        int firstMenRow = rowIndex + 2;
        int lastMenRow = rowIndex + 1 + params.getMenTeams().size();
        for (String team : params.getMenTeams()) {
            final String address = createDataRow(++rowIndex, team, firstMenRow, lastMenRow, true);
            addresses.putMen(team, address);
        }
        createDataRow(++rowIndex, null, 0, 0, true);

        int firstWomenRow = rowIndex + 2;
        int lastWomenRow = rowIndex + 1 + params.getWomenTeams().size();
        for (String team : params.getWomenTeams()) {
            final String address = createDataRow(++rowIndex, team, firstWomenRow, lastWomenRow, false);
            addresses.putWomen(team, address);
        }
        return addresses;
    }

    protected abstract String createDataRow(int rowIndex, String team, int firstRow, int lastRow, boolean isMan);

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

    protected void generateTitleCell(int tableRowIndex, int widthTable) {
        int rowIndex = tableRowIndex - 2;
        if (rowIndex < 0) {
            return;
        }
        final int cellIndex = 1;
        final Row row = sheet.createRow(rowIndex);
        row.setHeight((short) 600);
        final Cell cell = row.createCell(cellIndex);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 20);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(getSheetName());
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, cellIndex, cellIndex + widthTable));
    }
}
