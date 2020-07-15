package by.matskevich.protocol.service;

import by.matskevich.protocol.builders.*;
import by.matskevich.protocol.model.AllPlaces;
import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreationService {

    public Workbook generateProtocol(InputParams params) {
        Workbook wb = new XSSFWorkbook();
        int indexSheet = -1;

        CyclingBuilder cyclingBuilder = new CyclingBuilder(wb, params, ++indexSheet);
        OrientationBuilder orientationBuilder = new OrientationBuilder(wb, params, ++indexSheet);
        KTMBuilder ktmBuilder = new KTMBuilder(wb, params, ++indexSheet);
        WaterBuilder waterBuilder = new WaterBuilder(wb, params, ++indexSheet);

        AllPlaces allPlaces = new AllPlaces();
        allPlaces.setCyclingPlaces(cyclingBuilder.build());
        allPlaces.setKtmPlaces(ktmBuilder.build());
        allPlaces.setWaterPlaces(waterBuilder.build());
        allPlaces.setOrientationPlaces(orientationBuilder.build());


        FinalBuilder finalBuilder = new FinalBuilder(wb, params, ++indexSheet, allPlaces);
        finalBuilder.build();

        return wb;
    }
}
