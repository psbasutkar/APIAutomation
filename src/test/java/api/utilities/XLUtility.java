package api.utilities;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class XLUtility {

	    private String filePath;
	    private XSSFWorkbook workbook = new XSSFWorkbook();
	    private Sheet sheet;

	    public void ExcelUtil(String filePath, String sheetName) {
	        this.filePath = filePath;
	        try {
	            FileInputStream fis = new FileInputStream(filePath);
	            workbook = new XSSFWorkbook(fis);
	            sheet = workbook.getSheet(sheetName);
	            if (sheet == null) {
	                throw new RuntimeException("Sheet " + sheetName + " not found.");
	            }
	        } catch (IOException e) {
	            throw new RuntimeException("Could not read the Excel file: " + e.getMessage());
	        }
	    }

	    public String getCellData(int rowNum, int colNum) {
	        Row row = sheet.getRow(rowNum);
	        if (row == null) return "";
	        Cell cell = row.getCell(colNum);
	        if (cell == null) return "";
	        return getCellValueAsString(cell);
	    }

	    public int getRowCount() {
	        return sheet.getPhysicalNumberOfRows();
	    }

	    public int getColumnCount(int rowNum) {
	        Row row = sheet.getRow(rowNum);
	        return (row == null) ? 0 : row.getLastCellNum();
	    }

	    public Map<String, String> getRowDataAsMap(int rowNum) {
	        Map<String, String> rowData = new HashMap<>();
	        Row headerRow = sheet.getRow(0);
	        Row dataRow = sheet.getRow(rowNum);

	        if (headerRow == null || dataRow == null) return rowData;

	        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
	            String key = getCellValueAsString(headerRow.getCell(i));
	            String value = getCellValueAsString(dataRow.getCell(i));
	            rowData.put(key, value);
	        }

	        return rowData;
	    }

	    public void setCellData(int rowNum, int colNum, String value) {
	        try {
	            Row row = sheet.getRow(rowNum);
	            if (row == null) row = sheet.createRow(rowNum);
	            Cell cell = row.getCell(colNum);
	            if (cell == null) cell = row.createCell(colNum);
	            cell.setCellValue(value);

	            FileOutputStream fos = new FileOutputStream(filePath);
	            workbook.write(fos);
	            fos.close();
	        } catch (IOException e) {
	            throw new RuntimeException("Could not write to Excel file: " + e.getMessage());
	        }
	    }

	    private String getCellValueAsString(Cell cell) {
	        if (cell == null) return "";
	        switch (cell.getCellType()) {
	            case STRING: return cell.getStringCellValue();
	            case NUMERIC:
	                if (DateUtil.isCellDateFormatted(cell)) {
	                    return cell.getDateCellValue().toString();
	                } else {
	                    return String.valueOf(cell.getNumericCellValue());
	                }
	            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
	            case FORMULA: return cell.getCellFormula();
	            case BLANK: return "";
	            default: return "";
	        }
	    }

	    public void closeWorkbook() {
	        try {
	            if (workbook != null) {
	                workbook.close();
	            }
	        } catch (IOException e) {
	            System.out.println("Could not close workbook: " + e.getMessage());
	        }
	    }
	}


