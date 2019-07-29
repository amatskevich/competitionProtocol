package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

public class WaterBuilder extends BasisBuilder {

    static final String SHEET_NAME = "Вода";
    private static final int INITIAL_CELL_INDEX = 1;
    private static final int INITIAL_ROW_INDEX = 1;

    public WaterBuilder(Workbook wb, InputParams params, int indexSheet) {
        super(wb, params, indexSheet);
    }

    @Override
    String getSheetName() {
        return SHEET_NAME;
    }

    @Override
    public void build() {

        int rowIndex = INITIAL_ROW_INDEX;
        rowIndex = generateHeader(rowIndex);
        generateData(rowIndex);
        for (int i = INITIAL_CELL_INDEX; i < INITIAL_CELL_INDEX + 5; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private int generateHeader(int rowIndex) {
        final Row row1 = sheet.createRow(rowIndex);
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell teamCell = row1.createCell(cellIndex);
        teamCell.setCellValue("Цех");
        teamCell.setCellStyle(headerStyle);

        final Cell timeStartCell = row1.createCell(++cellIndex);
        timeStartCell.setCellValue("Время");
        timeStartCell.setCellStyle(headerStyle);

        final Cell fineTimeCell = row1.createCell(++cellIndex);
        fineTimeCell.setCellValue("Штрафное время");
        fineTimeCell.setCellStyle(headerStyle);

        final Cell sumTimeCell = row1.createCell(++cellIndex);
        sumTimeCell.setCellValue("Итоговое время");
        sumTimeCell.setCellStyle(headerStyle);

        final Cell placeCell = row1.createCell(++cellIndex);
        placeCell.setCellValue("Место");
        placeCell.setCellStyle(headerStyle);

        return rowIndex;
    }

    private void generateData(int rowIndex) {

        int firstMenRow = rowIndex + 2;
        int lastMenRow = rowIndex + 1 + params.getMenTeams().size();
        for (String team : params.getMenTeams()) {
            createDataRow(++rowIndex, team, firstMenRow, lastMenRow);
        }
        createDataRow(++rowIndex, null, 0, 0);

        int firstWomenRow = rowIndex + 2;
        int lastWomenRow = rowIndex + 1 + params.getWomenTeams().size();
        for (String team : params.getWomenTeams()) {
            createDataRow(++rowIndex, team, firstWomenRow, lastWomenRow);
        }
    }

    private void createDataRow(int rowIndex, String team, int firstRow, int lastRow) {
        final Row row = sheet.createRow(rowIndex);
        if (team == null) {
            return;
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
                + teamCell.getAddress().formatAsString()
                + "),\"\","
                + timeCell.getAddress().formatAsString()
                + "+"
                + fineCell.getAddress().formatAsString()
                + "/86400)";
        sumTimeCell.setCellFormula(formula);
        sumTimeCell.setCellStyle(timeStyle);

        final Cell placeCell = row.createCell(++cellIndex);
        final String colLetter = CellReference.convertNumToColString(sumTimeCell.getColumnIndex());
        String placeFormula = "IF(ISBLANK(" + teamCell.getAddress().formatAsString()
                + "),\"\","
                + "RANK(" + sumTimeCell.getAddress().formatAsString() + ", " + colLetter + firstRow + ':' + colLetter + lastRow + ",1))";
        placeCell.setCellFormula(placeFormula);
        placeCell.setCellStyle(dataStyle);
    }

}
