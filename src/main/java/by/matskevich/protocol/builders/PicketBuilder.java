package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.model.PlacesModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

abstract class PicketBuilder extends BasisBuilder {

    public PicketBuilder(Workbook wb, InputParams params, int indexSheet) {
        super(wb, params, indexSheet);
    }

    protected abstract int getInitialCellIndex();
    protected abstract int getInitialRowIndex();
    protected abstract int getCountPickets();

    @Override
    public PlacesModel build() {

        int rowIndex = getInitialRowIndex();

        rowIndex = generateHeader(rowIndex);

        final PlacesModel placesModel = generateData(rowIndex);

        for (int i = getInitialCellIndex(); i < getInitialCellIndex() + getCountPickets() + 6; i++) {
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
        int cellIndex = getInitialCellIndex();

        final Cell cell = row.createCell(cellIndex);
        cell.setCellValue(team);
        cell.setCellStyle(dataStyle);

        final Cell cell1 = row.createCell(++cellIndex);
        cell1.setCellStyle(timeStyle);

        final int firstPicketIndex = ++cellIndex;
        final int lastPicketIndex = firstPicketIndex + getCountPickets() - 1;
        for (int indexPicket = 1; cellIndex <= lastPicketIndex; ++cellIndex, ++indexPicket) {
            final Cell picketCell = row.createCell(cellIndex);
            picketCell.setCellStyle(dataStyle);
        }

        final Cell cell3 = row.createCell(cellIndex);
        cell3.setCellStyle(timeStyle);

        final Cell cell4 = row.createCell(++cellIndex);
//        cell4.setCellValue("Итоговое время");
//        cell4.setCellType(CellType.FORMULA);
        final String startTimeAddress = cell1.getAddress().formatAsString();
        final String finishTimeAddress = cell3.getAddress().formatAsString();
        String timeFormula= "IF(ISBLANK(" + startTimeAddress + "),\"\",IF(ISBLANK(" + finishTimeAddress + "),\"\"," + finishTimeAddress + "-" + startTimeAddress + "))";
        cell4.setCellFormula(timeFormula);
        cell4.setCellStyle(timeStyle);

        final Cell cell5 = row.createCell(++cellIndex);
//        cell5.setCellValue("Количество пикетов");
//        cell5.setCellType(CellType.FORMULA);
        final String shiftAddress = cell.getAddress().formatAsString();
        final String firstPicketAddress = row.getCell(firstPicketIndex).getAddress().formatAsString();
        final String lastPicketAddress = row.getCell(lastPicketIndex).getAddress().formatAsString();
        String picketsFormula= "IF(ISBLANK(" + shiftAddress + "),\"\",COUNTA(" + firstPicketAddress + ":" + lastPicketAddress + "))";
        cell5.setCellFormula(picketsFormula);
        cell5.setCellStyle(dataStyle);

        final Cell cell6 = row.createCell(++cellIndex);
//        cell6.setCellValue("Место");
//        cell5.setCellType(CellType.FORMULA);
//        String placeFormula= "RANK(" + cell6.getAddress().formatAsString() + ", K2:K10, desc))";//todo
//        cell6.setCellFormula(placeFormula);
        cell6.setCellStyle(dataStyle);
        return cell6.getAddress().formatAsString();
    }

    private int generateHeader(int headerRowIndex1) {
        final int headerRowIndex2 = headerRowIndex1 + 1;
        final Row row1 = sheet.createRow(headerRowIndex1);
        final Row row2 = sheet.createRow(headerRowIndex2);
        int cellIndex = getInitialCellIndex();

        final Cell cell = row1.createCell(cellIndex);
        cell.setCellValue("Цех");
        cell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell1 = row1.createCell(++cellIndex);
        cell1.setCellValue("Время старта");
        cell1.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell2 = row1.createCell(++cellIndex);
        cell2.setCellValue("№ пикета");
        cell2.setCellStyle(headerStyle);
        int lastPicketIndex = cellIndex + getCountPickets() - 1;
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex1, cellIndex, lastPicketIndex));
        for (int indexPicket = 1; cellIndex <= lastPicketIndex; ++cellIndex, ++indexPicket) {
            final Cell picketCell = row2.createCell(cellIndex);
            picketCell.setCellStyle(headerStyle);
            picketCell.setCellValue(indexPicket);
        }

        final Cell cell3 = row1.createCell(cellIndex);
        cell3.setCellValue("Время финиша");
        cell3.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell4 = row1.createCell(++cellIndex);
        cell4.setCellValue("Итоговое время");
        cell4.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell5 = row1.createCell(++cellIndex);
        cell5.setCellValue("Количество пикетов");
        cell5.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        final Cell cell6 = row1.createCell(++cellIndex);
        cell6.setCellValue("Место");
        cell6.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(headerRowIndex1, headerRowIndex2, cellIndex, cellIndex));

        return headerRowIndex2;
    }
}
