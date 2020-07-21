package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.model.PlacesModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

public class WaterBuilder extends BasisBuilder {

    public WaterBuilder(Workbook wb, String name, InputParams params, int indexSheet) {
        super(wb, name, params, indexSheet);
    }

    @Override
    public PlacesModel build() {

        generateTitleCell(INITIAL_ROW_INDEX, 5);
        generateHeader();
        final PlacesModel placesModel = generateData(INITIAL_ROW_INDEX);
        for (int i = INITIAL_CELL_INDEX; i < INITIAL_CELL_INDEX + 5; i++) {
            sheet.autoSizeColumn(i, true);
        }
        return placesModel;
    }

    private void generateHeader() {
        final Row row1 = sheet.createRow(INITIAL_ROW_INDEX);
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell teamCell = row1.createCell(cellIndex);
        teamCell.setCellValue("Цех");
        teamCell.setCellStyle(headerStyle);

        final Cell timeStartCell = row1.createCell(++cellIndex);
        timeStartCell.setCellValue("Время");
        timeStartCell.setCellStyle(headerStyle);

        final Cell fineTimeCell = row1.createCell(++cellIndex);
        fineTimeCell.setCellValue("Штрафное\r\nвремя");
        fineTimeCell.setCellStyle(headerStyle);

        final Cell sumTimeCell = row1.createCell(++cellIndex);
        sumTimeCell.setCellValue("Итоговое время");
        sumTimeCell.setCellStyle(headerStyle);

        final Cell placeCell = row1.createCell(++cellIndex);
        placeCell.setCellValue("Место");
        placeCell.setCellStyle(headerStyle);
    }

    @Override
    protected String createDataRow(int rowIndex, String team, int firstRow, int lastRow, boolean isMan) {
        final Row row = sheet.createRow(rowIndex);
        if (team == null) {
            return null;
        }
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell teamCell = row.createCell(cellIndex);
        teamCell.setCellValue(team);
        teamCell.setCellStyle(dataStyle);

        final Cell timeCell = row.createCell(++cellIndex);
        timeCell.setCellStyle(timeStyle);

        final Cell fineCell = row.createCell(++cellIndex);
        fineCell.setCellStyle(dataStyle);

        //=IF(ISBLANK(B8),"",C8+D8/86400)
        final Cell sumTimeCell = row.createCell(++cellIndex);
        String formula = "IF(ISBLANK("
                + timeCell.getAddress().formatAsString()
                + "),\"\","
                + timeCell.getAddress().formatAsString()
                + "+"
                + fineCell.getAddress().formatAsString()
                + "/86400)";
        sumTimeCell.setCellFormula(formula);
        sumTimeCell.setCellStyle(timeStyle);

        final Cell placeCell = row.createCell(++cellIndex);
        final String colLetter = CellReference.convertNumToColString(sumTimeCell.getColumnIndex());
        String placeFormula = "IF(ISBLANK(" + timeCell.getAddress().formatAsString()
                + "),\"\","
                + "RANK(" + sumTimeCell.getAddress().formatAsString() + ", " + colLetter + firstRow + ':' + colLetter + lastRow + ",1))";
        placeCell.setCellFormula(placeFormula);
        placeCell.setCellStyle(dataStyle);
        return placeCell.getAddress().formatAsString();
    }

}
