package threepm;

import java.io.File;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;

import com.opencsv.CSVWriter;


/**
 * @author the_fegati
 * @author kevmurumba
 */
public class ExcelReader {

    private final File file;
    DataFormatter formatter = new DataFormatter();

    public ExcelReader(File file) {
        this.file = file;
    }

    //Convert the excel into a List
    public void getRowAsListFromExcel() {

        //Array List to store the excel data
        List<String[]> csvList = new ArrayList<>();
        FileInputStream fis;
        Workbook workbook = null;
        int maxDataCount = 0;
        try {
            String fileExtension = file.toString().substring(file.toString().indexOf("."));
            OPCPackage oPCPackage = OPCPackage.open(file);
            //use xssf for xlsx format else hssf for xls format
            switch (fileExtension) {
                case ".xlsx":
                    workbook = new XSSFWorkbook(oPCPackage);
                    break;
                case ".xls":
//                    workbook = new HSSFWorkbook(new POIFSFileSystem(fis));
                    System.err.println("Wrong file type selected!");
                    break;
                default:
                    System.err.println("Wrong file type selected!");
                    break;
            }

            //get number of worksheets in the workbook
            int numberOfSheets = workbook.getNumberOfSheets();

            orgUnits orgUnits=new orgUnits();

            //iterating over each workbook sheet
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                int numberOfRows = sheet.getLastRowNum();

                for (int row = 4; row <= numberOfRows; row++) {
                    Row currentRow = sheet.getRow(row);
                    int numberOfCells = currentRow.getLastCellNum();

                    String facility = currentRow.getCell(0, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                    int mflcode = (int) currentRow.getCell(1, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                    int period = (int) currentRow.getCell(2, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();

                    String[] dataRows = new String[9];
                    HtsUids dataElements = new HtsUids();
                    int dataElementsLength = dataElements.IPD().length;
                    dataRows[6] = dataElements.IPD()[24][0];
                    dataRows[8] = dataElements.IPD()[25][0];
                    dataRows[4] = orgUnits.getOrgUid(mflcode);
                    for (int x = 3; x < dataElementsLength-2; x++) {
                        for (int y = 0; y < 2; y++) {
                            dataRows[0] = dataElements.IPD()[x][y];
                            System.out.printf("%s  ", dataRows[0]);
                        }
                        int dataValue = (int) currentRow.getCell(x, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        dataRows[2] = String.valueOf(period);
                        dataRows[3] = facility;
                        dataRows[7] = String.valueOf(dataValue);
                        csvList.add(dataRows);
                        System.out.printf("%d\t%s\t%d\t%s\t%s\t%d\t%s\n", period, facility, mflcode,dataRows[4], dataRows[6], dataValue, dataRows[8]);
                        System.out.printf("\n");
                    }

                    System.out.printf(String.valueOf(orgUnits.organizationalUnit()));
                    System.exit(0);
                }
            }
            workbook.close();
            writeRowToCSVFile(csvList);
        } catch (IOException | InvalidFormatException e) {
        }
    }

    /*
     * Write the rows into the CSV file
	 */
    private static void writeRowToCSVFile(List<String[]> cleanRows)
            throws IOException {
        File newFile = new File("/home/iita/Kelvin/JavaProjects/3pm-dataimport/output/oct17_jan18_ucsf_ipd_results_data.csv");
        try (CSVWriter csvWriter = new CSVWriter(new FileWriter(newFile))) {
            csvWriter.writeAll(cleanRows);
        }
    }

}
