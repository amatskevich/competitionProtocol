package by.matskevich.protocol.builders;

import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.Workbook;

public class OrientationBuilder extends PicketBuilder {

    static final String SHEET_NAME = "Ориентирование";
    private static final int INITIAL_CELL_INDEX = 1;
    private static final int INITIAL_ROW_INDEX = 1;

    public OrientationBuilder(Workbook wb, InputParams params, int indexSheet) {
        super(wb, params, indexSheet);
    }

    @Override
    public String getSheetName() {
        return SHEET_NAME;
    }

    @Override
    protected int getInitialCellIndex() {
        return INITIAL_CELL_INDEX;
    }

    @Override
    protected int getInitialRowIndex() {
        return INITIAL_ROW_INDEX;
    }

    @Override
    protected int getCountPickets() {
        return params.getCountOrientationPickets();
    }
}
