package by.matskevich.protocol.model;

public class AllPlaces {
    private PlacesModel cyclingPlaces;
    private PlacesModel ktmPlaces;
    private PlacesModel waterPlaces;
    private PlacesModel orientationPlaces;

    public PlacesModel getCyclingPlaces() {
        return cyclingPlaces;
    }

    public void setCyclingPlaces(PlacesModel cyclingPlaces) {
        this.cyclingPlaces = cyclingPlaces;
    }

    public PlacesModel getKtmPlaces() {
        return ktmPlaces;
    }

    public void setKtmPlaces(PlacesModel ktmPlaces) {
        this.ktmPlaces = ktmPlaces;
    }

    public PlacesModel getWaterPlaces() {
        return waterPlaces;
    }

    public void setWaterPlaces(PlacesModel waterPlaces) {
        this.waterPlaces = waterPlaces;
    }

    public PlacesModel getOrientationPlaces() {
        return orientationPlaces;
    }

    public void setOrientationPlaces(PlacesModel orientationPlaces) {
        this.orientationPlaces = orientationPlaces;
    }
}
