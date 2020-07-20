package by.matskevich.protocol.service;

import by.matskevich.protocol.builders.*;
import by.matskevich.protocol.model.*;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.LinkedList;
import java.util.List;

public class CreationService {

    public Workbook generateProtocol(InputParams params) {
        Workbook wb = new XSSFWorkbook();
        int indexSheet = -1;

        List<PlacesModel> allPlaces = new LinkedList<>();
        for (OrientationContest contest : params.getOrientationContests()) {
            PicketBuilder orientationBuilder = new PicketBuilder(wb, contest, params, ++indexSheet);
            allPlaces.add(orientationBuilder.build());
        }
        for (KtmContest ktmContest : params.getKtmContests()) {
            KTMBuilder ktmBuilder = new KTMBuilder(wb, ktmContest, params, ++indexSheet);
            allPlaces.add(ktmBuilder.build());
        }
        for (String name : params.getWaterContest()) {
            WaterBuilder waterBuilder = new WaterBuilder(wb, name, params, ++indexSheet);
            allPlaces.add(waterBuilder.build());
        }

        FinalBuilder finalBuilder = new FinalBuilder(wb, params, ++indexSheet, allPlaces);
        finalBuilder.build();

        return wb;
    }
}
