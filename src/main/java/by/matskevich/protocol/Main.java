package by.matskevich.protocol;

import by.matskevich.protocol.model.InputParams;
import by.matskevich.protocol.service.CreationService;

import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedList;

public class Main {

    public static void main(String[] args) throws IOException {
        System.out.println("Start program.");
        CreationService service = new CreationService();
        service.generateProtocol(prepareParams());
        System.out.println("Finish program.");
    }

    private static InputParams prepareParams() {
        InputParams params = new InputParams();
        params.setMenTeams(Arrays.asList("102", "401", "207"));
        params.setWomenTeams(Arrays.asList("604", "600"));
        params.setCountCyclePickets(5);
        params.setCountOrientationPickets(8);
        LinkedList<String> ktmPhases = new LinkedList<>();
        ktmPhases.add("веревки");
        ktmPhases.add("мышеловка");
        ktmPhases.add("гать");
        ktmPhases.add("кочки");
        ktmPhases.add("костер");
        ktmPhases.add("бревно");
        params.setKtmPhases(ktmPhases);
        LinkedHashMap<String, Double> stages = new LinkedHashMap<>();
        stages.put("Волейбол", 0.5);
        stages.put("Петанг", 0.5);
        stages.put("Художка", 1.0);
        params.setAdditionalStages(stages);
        return params;
    }
}
