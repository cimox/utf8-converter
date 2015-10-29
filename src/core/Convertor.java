package core;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

public abstract class Convertor implements Runnable {
    public static int convertExcel2CSV(File inputFile, File outputFile, boolean showCSVdata, boolean debugMode) throws IOException {        
        
    	FileInputStream fis;
      
        try {
            fis = new FileInputStream(inputFile);
        } catch (Exception e) {
            System.out.println("Error opening input xls file: "+inputFile+"\n");
            if (debugMode) e.printStackTrace();
            return -1;
        }

//        FileOutputStream fos;
        BufferedWriter out;
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outputFile), "UTF8"));         
        } catch (Exception e) {
            System.out.println("Error creating output csv file: "+outputFile+"\n");
            if (debugMode) e.printStackTrace();
            fis.close();
            return -1;
        }
        
        try {
            StringBuilder data = new StringBuilder();
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            if (debugMode == true) {
                System.out.println("Note:coordinates starts from 0,0 (A1)");
                System.out.println("getTopRow: " + sheet.getTopRow());
                System.out.println("getFirstRowNum: " + sheet.getFirstRowNum());
                System.out.println("getLastRowNum: " + sheet.getLastRowNum());
                System.out.println("getFirstCellNum: " + sheet.getRow(sheet.getFirstRowNum()).getFirstCellNum());
                System.out.println("getLastCellNum: " + sheet.getRow(sheet.getFirstRowNum()).getLastCellNum()+ " Note:Gets the index of the last cell contained in this row PLUS ONE.");
                System.out.println();
            }

            if (showCSVdata) System.out.println("CSV content:");

            Row firstRow = sheet.getRow(sheet.getFirstRowNum());            
            for (int iRow = sheet.getFirstRowNum(); iRow <= sheet.getLastRowNum(); iRow++) {

                Row row = sheet.getRow(iRow);
                sheet.getLastRowNum();

                for (int iCol = firstRow.getFirstCellNum(); iCol < firstRow.getLastCellNum() /*Gets the index of the last cell contained in this row PLUS ONE.*/; iCol++) {

                    Cell cell = row.getCell(iCol);

                    if (cell!=null) {
                        
                        switch (cell.getCellType()) {                            
                            case Cell.CELL_TYPE_BOOLEAN:
                                if (showCSVdata) System.out.print(cell.getBooleanCellValue()+";");
                                data.append(cell.getBooleanCellValue()).append(";");
                                break;
                            case Cell.CELL_TYPE_NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                            if (showCSVdata) System.out.print(cell.getDateCellValue()+";");
                                            data.append(cell.getDateCellValue()).append(";");
                                } 
                                else {
                                	if (showCSVdata) System.out.print(cell.getNumericCellValue()+";");
                                    data.append(cell.getNumericCellValue()).append(";");
                                }                                
                                break;
                            case Cell.CELL_TYPE_STRING:
                                if (showCSVdata) System.out.print(cell.getStringCellValue() + ";");
                                data.append(cell.getStringCellValue()).append(";");
                                break;
                            case Cell.CELL_TYPE_BLANK:
                                if (showCSVdata) System.out.print(";");
                                data.append(";");
                                break;
                            case Cell.CELL_TYPE_ERROR:
                                if (showCSVdata) System.out.print(cell.getErrorCellValue()+";");
                                data.append(";");
                                break;
                            case Cell.CELL_TYPE_FORMULA:
                                switch (cell.getCachedFormulaResultType()) {
                                    
                                    case Cell.CELL_TYPE_BOOLEAN:
                                        if (showCSVdata) System.out.print(cell.getBooleanCellValue()+";");
                                        data.append(cell.getBooleanCellValue()).append(";");
                                        break;
                                    case Cell.CELL_TYPE_NUMERIC:
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                                    if (showCSVdata) System.out.print(cell.getDateCellValue()+";");
                                                    data.append(cell.getDateCellValue()).append(";");
                                        } else {
                                            if (showCSVdata) System.out.print(cell.getNumericCellValue()+";");
                                            data.append(cell.getNumericCellValue()).append(";");
                                        }                                
                                        break;
                                    case Cell.CELL_TYPE_STRING:
                                        if (showCSVdata) System.out.print(cell.getStringCellValue()+";");
                                        data.append(cell.getStringCellValue()).append(";");
                                        break;
                                    case Cell.CELL_TYPE_BLANK:
                                        if (showCSVdata) System.out.print(";");
                                        data.append(";");
                                        break;
                                    case Cell.CELL_TYPE_ERROR:
                                        if (showCSVdata) System.out.print(cell.getErrorCellValue()+";");
                                        data.append(";");
                                        break;
                                    default:
                                        if (showCSVdata) System.out.print(cell.getStringCellValue()+";");
                                        data.append(";");
                                        break;
                                }
                                break;
                            default:
                                if (showCSVdata) System.out.print(cell.getStringCellValue()+";");
                                data.append(";");
                                break;
                        }
                    }
                    else {
                        if (showCSVdata) System.out.print(";");
                        data.append(";");
                    }
                    
                }
                if (showCSVdata) System.out.println();
                data.append('\n');
            }

//            fos.write(data.toString().getBytes());   
            String str = data.toString();
            out.write(data.toString());
        } catch (Exception e) {
            if (debugMode) e.printStackTrace();
            return -1;
        } finally {
            fis.close();
//            fos.close();
            out.close();
        }
        
        if (showCSVdata) System.out.println();
        return 1;
    }
}