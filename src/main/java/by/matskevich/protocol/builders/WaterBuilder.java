package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WaterBuilder extends BasisBuilder{

    private static final String SHEET_NAME = "Вода";
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

    }

}
