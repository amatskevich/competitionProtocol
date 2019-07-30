package by.matskevich.protocol.model;

import java.util.HashMap;
import java.util.Map;

public class PlacesModel {
    private final String sheetAddress;
    private Map<String, String> men = new HashMap<>();
    private Map<String, String> women = new HashMap<>();

    public PlacesModel(String sheetAddress) {
        this.sheetAddress = sheetAddress;
    }

    public String getManPlaceAddress(String team) {
        return men.get(team);
    }

    public String getWomanPlaceAddress(String team) {
        return women.get(team);
    }

    public void putMen(String team, String address) {
        men.put(team, address);
    }

    public void putWomen(String team, String address) {
        women.put(team, address);
    }

    public String getSheetAddress() {
        return sheetAddress;
    }

    public Map<String, String> getMen() {
        return men;
    }

    public Map<String, String> getWomen() {
        return women;
    }
}
