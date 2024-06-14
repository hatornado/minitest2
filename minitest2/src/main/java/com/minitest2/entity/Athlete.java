package com.minitest2.entity;

public class Athlete {
    private String name;
    private String unit;
    private int birthYear;

    public Athlete(String name, String unit, int birthYear) {
        this.name = name;
        this.unit = unit;
        this.birthYear = birthYear;
    }

    @Override
    public String toString() {
        return String.format("%s - %s - %d", name, unit, birthYear);
    }
}
