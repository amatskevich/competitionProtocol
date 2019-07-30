package by.matskevich.protocol.service;

import by.matskevich.protocol.builders.*;
import by.matskevich.protocol.model.AllPlaces;
import by.matskevich.protocol.model.InputParams;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CreationService {

    public void generateProtocol(InputParams params) throws IOException {

        FileOutputStream out = new FileOutputStream("workbook.xlsx");
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

// write the workbook to the output stream
// close our file (don't blow out our file handles
        wb.write(out);
        out.close();

    }
}
