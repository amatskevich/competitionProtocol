package by.matskevich.protocol.model;

import java.util.LinkedList;
import java.util.List;

public class InputParams {

    private List<String> menTeams;
    private List<String> womenTeams;
    private List<OrientationContest> orientationContests;
    private List<KtmContest> ktmContests;
    private List<String> waterContest;
    private LinkedList<SimpleContest> simpleContests;

    public List<String> getMenTeams() {
        return menTeams;
    }

    public void setMenTeams(List<String> menTeams) {
        this.menTeams = menTeams;
    }

    public List<String> getWomenTeams() {
        return womenTeams;
    }

    public void setWomenTeams(List<String> womenTeams) {
        this.womenTeams = womenTeams;
    }

    public List<OrientationContest> getOrientationContests() {
        return orientationContests;
    }

    public void setOrientationContests(List<OrientationContest> orientationContests) {
        this.orientationContests = orientationContests;
    }

    public List<KtmContest> getKtmContests() {
        return ktmContests;
    }

    public void setKtmContests(List<KtmContest> ktmContests) {
        this.ktmContests = ktmContests;
    }

    public List<String> getWaterContest() {
        return waterContest;
    }

    public void setWaterContest(List<String> waterContest) {
        this.waterContest = waterContest;
    }

    public LinkedList<SimpleContest> getSimpleContests() {
        return simpleContests;
    }

    public void setSimpleContests(LinkedList<SimpleContest> simpleContests) {
        this.simpleContests = simpleContests;
    }

    @Override
    public String toString() {
        return "InputParams{" +
                "menTeams=" + menTeams +
                ", womenTeams=" + womenTeams +
                ", orientationContests=" + orientationContests +
                ", ktmContests=" + ktmContests +
                ", waterContest=" + waterContest +
                ", simpleContests=" + simpleContests +
                '}';
    }
}
