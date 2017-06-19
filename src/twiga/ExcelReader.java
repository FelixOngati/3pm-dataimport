package twiga;

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
 *
 * @author the_fegati
 */
public class ExcelReader {

    private final File file;
    DataFormatter formatter = new DataFormatter();

    public ExcelReader(File file) {
        this.file = file;
    }

    public void getRowAsListFromExcel() {
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

            //iterating over each workbook sheet
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                int numberOfRows = sheet.getLastRowNum();

                for (int row = 3; row <= numberOfRows; row++) {
                    Row currentRow = sheet.getRow(row);
                    int numberOfCells = currentRow.getLastCellNum();
                    
                    int attributeOptionCombo = (int) currentRow.getCell(0, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                    int organisationunitid = (int) currentRow.getCell(1, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                    String period = currentRow.getCell(2, Row.CREATE_NULL_AS_BLANK).getStringCellValue();

                    for (int cell = 3; cell <= numberOfCells; cell++) {
                        int dataelementId = (int) sheet.getRow(1).getCell(cell, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        int categoryOptionComboId = (int) sheet.getRow(2).getCell(cell, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        int dataValue = (int) currentRow.getCell(cell, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        String[] dataRows = new String[6];
                        dataRows[0] = String.valueOf(dataelementId);
                        dataRows[1] = period;
                        dataRows[2] = String.valueOf(organisationunitid);
                        dataRows[3] = String.valueOf(categoryOptionComboId);
                        dataRows[4] = String.valueOf(attributeOptionCombo);
                        dataRows[5] = String.valueOf(dataValue);
                        csvList.add(dataRows);
//                        Pick from here after church  
                        System.out.printf("%d\t%d\t%s\t%d\t%d\t%d\n", attributeOptionCombo, organisationunitid, period, dataelementId, categoryOptionComboId, dataValue);
                    }
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
            File newFile = new File("/home/fegati/NetBeansProjects/Twiga/output/targets.csv");
        try (CSVWriter csvWriter = new CSVWriter(new FileWriter(newFile))) {
            csvWriter.writeAll(cleanRows);
        }
}

}
