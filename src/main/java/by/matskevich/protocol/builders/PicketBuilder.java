package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.model.OrientationContest;
import by.matskevich.protocol.model.PlacesModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

public class PicketBuilder extends BasisBuilder {

    private final OrientationContest contest;

    public PicketBuilder(Workbook wb, OrientationContest contest, InputParams params, int indexSheet) {
        super(wb, contest.getName(), params, indexSheet);
        this.contest = contest;
    }

    @Override
    public PlacesModel build() {

        int rowIndex = INITIAL_ROW_INDEX;
        generateTitleCell(INITIAL_ROW_INDEX, 6 + contest.getCountOfPickets());

        rowIndex = generateHeader(rowIndex);

        final PlacesModel placesModel = generateData(rowIndex);

        for (int i = INITIAL_CELL_INDEX; i < INITIAL_CELL_INDEX + contest.getCountOfPickets() + 6; i++) {
            sheet.autoSizeColumn(i, true);
        }

        return placesModel;
    }

    @Override
    protected String createDataRow(int indexRow, String team, int firstRow, int lastRow, boolean isMan) {
        final Row row = sheet.createRow(indexRow);
        if (team == null) {
            return null;
        }
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell teamCell = row.createCell(cellIndex);
        teamCell.setCellValue(team);
        teamCell.setCellStyle(dataStyle);

        final Cell startTimeCell = row.createCell(++cellIndex);
        startTimeCell.setCellStyle(timeStyle);

        final int firstPicketIndex = ++cellIndex;
        final int lastPicketIndex = firstPicketIndex + contest.getCountOfPickets() - 1;
        for (int indexPicket = 1; cellIndex <= lastPicketIndex; ++cellIndex, ++indexPicket) {
            final Cell picketCell = row.createCell(cellIndex);
            picketCell.setCellStyle(dataStyle);
        }

        final Cell finishTimeCell = row.createCell(cellIndex);
        finishTimeCell.setCellStyle(timeStyle);

        final Cell resultTimeCell = row.createCell(++cellIndex);
        final String startTimeAddress = startTimeCell.getAddress().formatAsString();
        final String finishTimeAddress = finishTimeCell.getAddress().formatAsString();
        String timeFormula = "IF(ISBLANK(" + startTimeAddress + "),\"\",IF(ISBLANK(" + finishTimeAddress + "),\"\"," + finishTimeAddress + "-" + startTimeAddress + "))";
        resultTimeCell.setCellFormula(timeFormula);
        resultTimeCell.setCellStyle(timeStyle);

        final Cell amountOfPicketsCell = row.createCell(++cellIndex);
        final String teamAddress = teamCell.getAddress().formatAsString();
        final String firstPicketAddress = row.getCell(firstPicketIndex).getAddress().formatAsString();
        final String lastPicketAddress = row.getCell(lastPicketIndex).getAddress().formatAsString();
        String picketsFormula = "IF(ISBLANK(" + teamAddress + "),\"\",COUNTA(" + firstPicketAddress + ":" + lastPicketAddress + "))";
        amountOfPicketsCell.setCellFormula(picketsFormula);
        amountOfPicketsCell.setCellStyle(dataStyle);

        //=COUNTIF(N4:N6,">"&N4)+1+SUMPRODUCT(--(N4:N6=N4),--(M4:M6<M4))
        final String timeColumn = CellReference.convertNumToColString(resultTimeCell.getColumnIndex());
        final String timeAddress = resultTimeCell.getAddress().formatAsString();
        final String picketsAddress = amountOfPicketsCell.getAddress().formatAsString();
        final String picketsColumn = CellReference.convertNumToColString(amountOfPicketsCell.getColumnIndex());
        final Cell placeCell = row.createCell(++cellIndex);
        String placeFormula = "IF(ISBLANK(" + finishTimeCell.getAddress().formatAsString() + "),\"\",COUNTIF("
                + picketsColumn + firstRow + ":" + picketsColumn + lastRow
                + ",\">\"&" + picketsAddress
                + ")+1+SUMPRODUCT(--("
                + picketsColumn + firstRow + ":" + picketsColumn + lastRow
                + "=" + picketsAddress
                + "),--("
                + timeColumn + firstRow + ":" + timeColumn + lastRow
                + "<" + timeAddress + ")))";
        placeCell.setCellFormula(placeFormula);
        placeCell.setCellStyle(dataStyle);
        return placeCell.getAddress().formatAsString();
    }

    private int generateHeader(int headerRowIndex1) {
        final int headerRowIndex2 = headerRowIndex1 + 1;
        final Row row1 = sheet.createRow(headerRowIndex1);
        final Row row2 = sheet.createRow(headerRowIndex2);
        int cellIndex = INITIAL_CELL_INDEX;

        final Cell cell = row1.createCell(cellIndex);
        cell.setCellValue("Цех");
        cell.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell1 = row1.createCell(++cellIndex);
        cell1.setCellValue("Время старта");
        cell1.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell2 = row1.createCell(++cellIndex);
        cell2.setCellValue("№ пикета");
        cell2.setCellStyle(headerStyle);
        int lastPicketIndex = cellIndex + contest.getCountOfPickets() - 1;
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex1, cellIndex, lastPicketIndex));
        for (int indexPicket = 1; cellIndex <= lastPicketIndex; ++cellIndex, ++indexPicket) {
            row1.createCell(cellIndex).setCellStyle(headerStyle);
            final Cell picketCell = row2.createCell(cellIndex);
            picketCell.setCellStyle(headerStyle);
            picketCell.setCellValue(indexPicket);
        }

        final Cell cell3 = row1.createCell(cellIndex);
        cell3.setCellValue("Время финиша");
        cell3.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell4 = row1.createCell(++cellIndex);
        cell4.setCellValue("Итоговое время");
        cell4.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell5 = row1.createCell(++cellIndex);
        cell5.setCellValue("Количество пикетов");
        cell5.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell6 = row1.createCell(++cellIndex);
        cell6.setCellValue("Место");
        cell6.setCellStyle(headerStyle);
        row2.createCell(cellIndex).setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        return headerRowIndex2;
    }
}
