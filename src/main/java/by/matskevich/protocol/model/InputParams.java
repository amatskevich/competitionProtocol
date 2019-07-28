package by.matskevich.protocol.model;

import java.util.LinkedList;
import java.util.List;

public class InputParams {

    private List<String> menTeams;
    private List<String> womenTeams;
    private Integer countCyclePickets;
    private Integer countOrientationPickets;
    private LinkedList<String> ktmPhases;

    public LinkedList<String> getKtmPhases() {
        return ktmPhases;
    }

    public void setKtmPhases(LinkedList<String> ktmPhases) {
        this.ktmPhases = ktmPhases;
    }

    public Integer getCountOrientationPickets() {
        return countOrientationPickets;
    }

    public void setCountOrientationPickets(Integer countOrientationPickets) {
        this.countOrientationPickets = countOrientationPickets;
    }

    public Integer getCountCyclePickets() {
        return countCyclePickets;
    }

    public void setCountCyclePickets(Integer countCyclePickets) {
        this.countCyclePickets = countCyclePickets;
    }

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

}
