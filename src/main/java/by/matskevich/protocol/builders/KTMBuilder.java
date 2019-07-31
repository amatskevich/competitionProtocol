package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.model.PlacesModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.ArrayList;
import java.util.List;

public class KTMBuilder extends BasisBuilder {

    static final String SHEET_NAME = "КТМ";
    private static final int INITIAL_CELL_INDEX = 1;
    private static final int INITIAL_ROW_INDEX = 3;

    public KTMBuilder(Workbook wb, InputParams params, int indexSheet) {
        super(wb, params, indexSheet);
    }

    @Override
    String getSheetName() {
        return SHEET_NAME;
    }

    @Override
    public PlacesModel build() {

        int rowIndex = INITIAL_ROW_INDEX;
        generateTitleCell(INITIAL_ROW_INDEX, 14);
        rowIndex = generateHeader(rowIndex);
        final PlacesModel placesModel = generateData(rowIndex);
        for (int i = INITIAL_CELL_INDEX; i < INITIAL_CELL_INDEX + 2 * params.getKtmPhases().size() + 5; i++) {
            sheet.autoSizeColumn(i, true);
        }
        return placesModel;
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

        final Cell startTimeCell = row.createCell(++cellIndex);
        startTimeCell.setCellStyle(timeStyle);

        List<String> fineAddresses = new ArrayList<>();
        List<String> pauseAddresses = new ArrayList<>();

        for (int i = 0; i < params.getKtmPhases().size(); ++i) {
            final Cell fineCell = row.createCell(++cellIndex);
            fineCell.setCellStyle(dataStyle);
            fineAddresses.add(fineCell.getAddress().formatAsString());
            final Cell pauseCell = row.createCell(++cellIndex);
            pauseCell.setCellStyle(dataStyle);
            pauseAddresses.add(pauseCell.getAddress().formatAsString());
        }

        final Cell finishTimeCell = row.createCell(++cellIndex);
        finishTimeCell.setCellStyle(timeStyle);

        //=IF(ISBLANK(R4),"",R4-C4+(-E4-G4-I4-K4-M4-O4-Q4+D4+F4+H4+J4+L4+N4+P4)/86400)
        final Cell sumTimeCell = row.createCell(++cellIndex);
        final String startTimeAddress = startTimeCell.getAddress().formatAsString();
        final String finishTimeAddress = finishTimeCell.getAddress().formatAsString();
        StringBuilder formula = new StringBuilder("IF(ISBLANK(");
        formula.append(finishTimeAddress);
        formula.append("),\"\",");
        formula.append(finishTimeAddress);
        formula.append('-');
        formula.append(startTimeAddress);
        formula.append("+(");
        formula.append(String.join("+", fineAddresses));
        formula.append('-');
        formula.append(String.join("-", pauseAddresses));
        formula.append(")/86400)");
        sumTimeCell.setCellFormula(formula.toString());
        sumTimeCell.setCellStyle(timeStyle);

        final Cell placeCell = row.createCell(++cellIndex);
        final String colLetter = CellReference.convertNumToColString(sumTimeCell.getColumnIndex());
        String placeFormula = "IF(ISBLANK(" + finishTimeAddress + "),\"\","
                + "RANK(" + sumTimeCell.getAddress().formatAsString() + ", " + colLetter + firstRow + ':' + colLetter + lastRow + ",1))";
        placeCell.setCellFormula(placeFormula);
        placeCell.setCellStyle(dataStyle);
        return placeCell.getAddress().formatAsString();
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

        final Cell timeStartCell = row1.createCell(++cellIndex);
        timeStartCell.setCellValue("Время старта");
        timeStartCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        for (String name : params.getKtmPhases()) {
            cellIndex = addPhase(name, row1, row2, cellIndex, headerRowIndex1);
        }

        final Cell timeFinishCell = row1.createCell(++cellIndex);
        timeFinishCell.setCellValue("Время финиша");
        timeFinishCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell sumTimeCell = row1.createCell(++cellIndex);
        sumTimeCell.setCellValue("Итоговое время");
        sumTimeCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell placeCell = row1.createCell(++cellIndex);
        placeCell.setCellValue("Место");
        placeCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        return headerRowIndex2;
    }

    private int addPhase(String name, Row row1, Row row2, int cellIndex, int headerRowIndex1) {
        final Cell cell = row1.createCell(++cellIndex);
        final int cellIndex2 = cellIndex + 1;
        cell.setCellValue("Этап " + name);
        cell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex1, cellIndex, cellIndex2));

        final Cell cellLeft = row2.createCell(cellIndex);
        cellLeft.setCellValue("Штрафное \n время");
        cellLeft.setCellStyle(headerStyle);

        final Cell cellRight = row2.createCell(cellIndex2);
        cellRight.setCellValue("Отсечка");
        cellRight.setCellStyle(headerStyle);
        return cellIndex2;
    }
}
