/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package twiga;

import java.io.File;

/**
 *
 * @author fegati
 */
public class Twiga {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        File excelFile = new File("/home/fegati/Documents/twiga/data/targetdata-processed.xlsx");
        ExcelReader er = new ExcelReader(excelFile);
        er.getRowAsListFromExcel();
    }
    
}
