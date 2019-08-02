package by.matskevich.protocol;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.service.CreationService;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.File;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
        System.out.println("Start program.");
        ObjectMapper objectMapper = new ObjectMapper();
        final InputParams params = objectMapper.readValue(new File("input.json"), InputParams.class);
        CreationService service = new CreationService();
        service.generateProtocol(params);
        System.out.println("Finish program.");
    }
}
