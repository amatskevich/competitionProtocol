package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.model.PlacesModel;
import by.matskevich.protocol.model.SimpleContest;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.List;

public class FinalBuilder extends BasisBuilder {

    private final List<PlacesModel> allPlaces;

    public FinalBuilder(Workbook wb, InputParams params, int indexSheet, List<PlacesModel> allPlaces) {
        super(wb, "Сводный протокол", params, indexSheet);
        this.allPlaces = allPlaces;
    }

    @Override
    public PlacesModel build() {

        int rowIndex = INITIAL_ROW_INDEX;
        generateTitleCell(INITIAL_ROW_INDEX, 12);
        rowIndex = generateHeader(rowIndex);
        generateData(rowIndex);
        int countOfColumns = INITIAL_CELL_INDEX + 2 * params.getSimpleContests().size() + allPlaces.size() + 3;
        for (int i = INITIAL_CELL_INDEX; i < countOfColumns; i++) {
            sheet.autoSizeColumn(i, true);
        }
        return null;
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
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        for (PlacesModel placesModel : allPlaces) {
            final Cell cell = row1.createCell(++cellIndex);
            cell.setCellValue(placesModel.getSheetAddress());
            cell.setCellStyle(headerStyle);
            row2.createCell(cellIndex).setCellStyle(headerStyle);
            sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));
        }

        for (SimpleContest simpleContest : params.getSimpleContests()) {
            cellIndex = addStageHeader(simpleContest.getName(), row1, row2, cellIndex, headerRowIndex1);
        }

        final Cell sumTimeCell = row1.createCell(++cellIndex);
        sumTimeCell.setCellValue("Итоговые\r\nбаллы");
        sumTimeCell.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell placeCell = row1.createCell(++cellIndex);
        placeCell.setCellValue("Итоговые\r\nместо");
        placeCell.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        return headerRowIndex2;
    }

    private int addStageHeader(String name, Row row1, Row row2, int cellIndex, int headerRowIndex1) {
        final Cell cell = row1.createCell(++cellIndex);
        final int cellIndex2 = cellIndex + 1;
        row1.createCell(cellIndex2).setCellStyle(headerStyle);
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

    @Override
    protected String createDataRow(int rowIndex, String team, int firstRow, int lastRow, boolean isMen) {

        final Row row = sheet.createRow(rowIndex);
        if (team == null) {
            return null;
        }
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell teamCell = row.createCell(cellIndex);
        teamCell.setCellValue(team);
        teamCell.setCellStyle(dataStyle);

        List<String> cellAddresses = new ArrayList<>();
        for (PlacesModel placesModel : allPlaces) {
            final Cell cell = row.createCell(++cellIndex);
            cell.setCellFormula(generateLinkAddress(placesModel.getSheetAddress(), placesModel, team, isMen));
            cell.setCellStyle(dataStyle);
            cellAddresses.add(cell.getAddress().formatAsString());
        }

        List<String> scoreAddresses = new ArrayList<>();

        for (SimpleContest simpleContest : params.getSimpleContests()) {
            final Cell stagePlaceCell = row.createCell(++cellIndex);
            stagePlaceCell.setCellStyle(dataStyle);
            final Cell scoreStageCell = row.createCell(++cellIndex);
            scoreStageCell.setCellStyle(dataStyle);
            scoreStageCell.setCellFormula(stagePlaceCell.getAddress().formatAsString() + "*" + simpleContest.getRate());
            scoreAddresses.add(scoreStageCell.getAddress().formatAsString());
        }

//        =C5+D5+E5+F5+H5+J5+K5
        final Cell finalScoreCell = row.createCell(++cellIndex);
        finalScoreCell.setCellStyle(dataStyle);
        cellAddresses.addAll(scoreAddresses);
        finalScoreCell.setCellFormula(String.join("+", cellAddresses));


        final Cell placeCell = row.createCell(++cellIndex);
        final String colLetter = CellReference.convertNumToColString(finalScoreCell.getColumnIndex());
        String placeFormula = "IF(ISBLANK(" + teamCell.getAddress().formatAsString() + "),\"\","
                + "RANK(" + finalScoreCell.getAddress().formatAsString() + ", " + colLetter + firstRow + ':' + colLetter + lastRow + ",1))";
        placeCell.setCellFormula(placeFormula);
        placeCell.setCellStyle(dataStyle);
        return null;
    }

    private String generateLinkAddress(String sheetName, PlacesModel places, String team, boolean isMen) {
        String address = isMen ? places.getManPlaceAddress(team) : places.getWomanPlaceAddress(team);
        return "'" + sheetName + "'!" + address;
    }
}
