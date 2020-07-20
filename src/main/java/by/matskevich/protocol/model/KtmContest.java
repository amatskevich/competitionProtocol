package by.matskevich.protocol.model;

import java.util.LinkedList;

public class KtmContest {
    private String name;
    LinkedList<String> phases;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public LinkedList<String> getPhases() {
        return phases;
    }

    public void setPhases(LinkedList<String> phases) {
        this.phases = phases;
    }
}
