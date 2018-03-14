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

                for (int row = 4; row <= numberOfRows; row++) {
                    Row currentRow = sheet.getRow(row);
                    int numberOfCells = currentRow.getLastCellNum();
                    
                    String facility = currentRow.getCell(0,Row.CREATE_NULL_AS_BLANK).getStringCellValue();
//                     System.out.println(facility);
//                    System.exit(0);
                    int mflcode = (int) currentRow.getCell(1,Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                   
                    int period = (int) currentRow.getCell(2,Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                    
                    for (int cell = 3; cell <= numberOfCells; cell++) {
                        String dataelementId = sheet.getRow(0).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        String dataelementname = sheet.getRow(1).getCell(cell,Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        String categoryOptionCombo = sheet.getRow(2).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        String attributeOptionCombo = sheet.getRow(3).getCell(cell, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
                        int dataValue = (int) currentRow.getCell(cell, Row.CREATE_NULL_AS_BLANK).getNumericCellValue();
                        if(dataValue  == 0)
                            continue;
                        String[] dataRows = new String[8];
                        dataRows[0] = String.valueOf(dataelementId);
                        dataRows[1] = dataelementname;
                        dataRows[2] = String.valueOf(period);
                        dataRows[3] = facility;
                        dataRows[4] = String.valueOf(mflcode);
                        dataRows[5] = categoryOptionCombo;
                        dataRows[6] = attributeOptionCombo;
                        dataRows[7] = String.valueOf(dataValue);
                        System.out.println(cell);
                        csvList.add(dataRows);
//                        Pick from here after church  
                        System.out.printf("%s\t%s\t%s\t%s\t%d\t%s\t%s\t%d\n", dataelementId,dataelementname,period,facility,mflcode,categoryOptionCombo,attributeOptionCombo,dataValue);
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
            File newFile = new File("/home/fegati/NetBeansProjects/Twiga/output/ucsf_c&t_oct-dec_results_data.csv");
        try (CSVWriter csvWriter = new CSVWriter(new FileWriter(newFile))) {
            csvWriter.writeAll(cleanRows);
        }
}

}
