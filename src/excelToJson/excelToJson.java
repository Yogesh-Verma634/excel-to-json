package excelToJson;

import org.apache.commons.lang.WordUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.*;
import java.util.Iterator;

	public class excelToJson {

		public static void main(String[] args) throws IOException, InvalidFormatException, JSONException {

	           InputStream inp = new FileInputStream("SomeFile.xlsx");

	            XSSFWorkbook wb = new XSSFWorkbook(inp);
	// Get the first Sheet.
	            XSSFSheet sheet=wb.getSheetAt(0);
	            XSSFRow excelrow;
	            XSSFCell cell;
	            Iterator rows = sheet.rowIterator();
	                // Start constructing JSON.
	            JSONObject json = new JSONObject();

	             // Iterate through the rows.
	            JSONArray jsonrows = new JSONArray();
	        while (rows.hasNext())
	        {
	                excelrow = (XSSFRow) rows.next();
	                JSONObject someId = new JSONObject();
	                Iterator excelcells = excelrow.cellIterator();
	                // Iterate through the cells.
	                JSONArray cells = new JSONArray();
	            while (excelcells.hasNext())
	            {
	                    cell = (XSSFCell) excelcells.next();
	                if(cell.getCellType() == cell.CELL_TYPE_NUMERIC){
	                	someId.put("Id", cell.getNumericCellValue());
	                }
	                else{
	                    if(cell.getStringCellValue().endsWith(" ")){
	                    someId.put("Title", cell.getStringCellValue().trim());
	                    }
	                    else{
	                    	someId.put("Title",cell.getStringCellValue());
	                }
	                }
	              //  jRow.put( "cell", cells );

	            }
	            jsonrows.put( someId );
	        }

	            // Create the JSON.
	            json.put( "someId", jsonrows );

	// Get the JSON text.
	           //  json.toString();
	             FileWriter file = new FileWriter("file1.json");
	     			file.write(json.toString());
	     			//System.out.println("Successfully Copied JSON Object to File...");
	     			//System.out.println("\nJSON Object: " + json);
	     			file.flush();
	     			file.close();


	}
	}
