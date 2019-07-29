package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class FinalBuilder extends BasisBuilder {

    private static final String SHEET_NAME = "Сводный протокол";
    private static final int INITIAL_CELL_INDEX = 1;
    private static final int INITIAL_ROW_INDEX = 1;

    public FinalBuilder(Workbook wb, InputParams params, int indexSheet) {
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
        for (int i = INITIAL_CELL_INDEX; i < INITIAL_CELL_INDEX + 2 * params.getAdditionalStages().size() + 7; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private int generateHeader(int headerRowIndex1) {

        final int headerRowIndex2 = headerRowIndex1 + 1;
        final Row row1 = sheet.createRow(headerRowIndex1);
        final Row row2 = sheet.createRow(headerRowIndex2);
        row2.setHeight((short) -1);
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell teamCell = row1.createCell(cellIndex);
        teamCell.setCellValue("Цех");
        teamCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cyclingCell = row1.createCell(++cellIndex);
        cyclingCell.setCellValue(CyclingBuilder.SHEET_NAME);
        cyclingCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell ktmCell = row1.createCell(++cellIndex);
        ktmCell.setCellValue(KTMBuilder.SHEET_NAME);
        ktmCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell waterCell = row1.createCell(++cellIndex);
        waterCell.setCellValue(WaterBuilder.SHEET_NAME);
        waterCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell orientationCell = row1.createCell(++cellIndex);
        orientationCell.setCellValue(OrientationBuilder.SHEET_NAME);
        orientationCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        for (String name : params.getAdditionalStages().keySet()) {
            cellIndex = addStageHeader(name, row1, row2, cellIndex, headerRowIndex1);
        }

        final Cell sumTimeCell = row1.createCell(++cellIndex);
        sumTimeCell.setCellValue("Итоговые \nбаллы");
        sumTimeCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell placeCell = row1.createCell(++cellIndex);
        placeCell.setCellValue("Итоговые \nместо");
        placeCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        return headerRowIndex2;
    }
    private int addStageHeader(String name, Row row1, Row row2, int cellIndex, int headerRowIndex1) {
        final Cell cell = row1.createCell(++cellIndex);
        final int cellIndex2 = cellIndex + 1;
        cell.setCellValue(name);
        cell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex1, cellIndex, cellIndex2));

        final Cell cellLeft = row2.createCell(cellIndex);
        cellLeft.setCellValue("Место");
        cellLeft.setCellStyle(headerStyle);

        final Cell cellRight = row2.createCell(cellIndex2);
        cellRight.setCellValue("Балл");
        cellRight.setCellStyle(headerStyle);
        return cellIndex2;
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

        final Cell cyclingCell = row.createCell(++cellIndex);
        cyclingCell.setCellStyle(dataStyle);

        final Cell ktmCell = row.createCell(++cellIndex);
        ktmCell.setCellStyle(dataStyle);

        final Cell waterCell = row.createCell(++cellIndex);
        waterCell.setCellStyle(dataStyle);

        final Cell orientationCell = row.createCell(++cellIndex);
        orientationCell.setCellStyle(dataStyle);

        List<String> scoreAddresses = new ArrayList<>();

        for (Map.Entry<String, Double> entry : params.getAdditionalStages().entrySet()) {
            final Cell stagePlaceCell = row.createCell(++cellIndex);
            stagePlaceCell.setCellStyle(dataStyle);
            final Cell scoreStageCell = row.createCell(++cellIndex);
            scoreStageCell.setCellStyle(dataStyle);
            scoreStageCell.setCellFormula(stagePlaceCell.getAddress().formatAsString() + "*" + entry.getValue());
            scoreAddresses.add(scoreStageCell.getAddress().formatAsString());
        }

//        =C5+D5+E5+F5+H5+J5+K5
        final Cell finalScoreCell = row.createCell(++cellIndex);
        finalScoreCell.setCellStyle(dataStyle);
        List<String> cellAddresses = new ArrayList<>();
        cellAddresses.add(cyclingCell.getAddress().formatAsString());
        cellAddresses.add(ktmCell.getAddress().formatAsString());
        cellAddresses.add(waterCell.getAddress().formatAsString());
        cellAddresses.add(orientationCell.getAddress().formatAsString());
        cellAddresses.addAll(scoreAddresses);
        finalScoreCell.setCellFormula(String.join("+", cellAddresses));


        final Cell placeCell = row.createCell(++cellIndex);
        final String colLetter = CellReference.convertNumToColString(finalScoreCell.getColumnIndex());
        String placeFormula = "IF(ISBLANK(" + teamCell.getAddress().formatAsString() + "),\"\","
                + "RANK(" + finalScoreCell.getAddress().formatAsString() + ", " + colLetter + firstRow + ':' + colLetter + lastRow + ",1))";
        placeCell.setCellFormula(placeFormula);
        placeCell.setCellStyle(dataStyle);
    }

}
