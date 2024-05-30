package org.example;


import java.io.IOException;

public class App
{
    public static void main( String[] args ) throws IOException {
        ParseAndCreateExcell.parseAndCreateExcell("data/Ug.xlsx", 4);
        ParseAndCreateExcell.parseAndCreateExcell("data/Manich.xlsx", 5);
        ParseAndCreateExcell.parseAndCreateExcell("data/Kuban.xlsx", 5);
        ParseAndCreateExcell.closeTable();


    }
}
