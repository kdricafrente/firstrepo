package testapp1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiDemo {


	private static final String wb = null;
	private static String ws;

	public static void main(String[] args) {
		
		createworkbook("employees", "records");
		
//       readexcel("employees", "records");\
		
//		appendrow("employees", "records", "4", "rina", "santos");

		} 
	
	public static void appendrow(String wbook, String wsheet, String id, String name, String department){
		
		
		try {
			FileInputStream file = new FileInputStream(new File(wb + ".xlsx"));
			
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet(ws);
			
			int rowlastnum = sheet.getLastRowNum();
			Row newrow = sheet.createRow(rowlastnum + 1);
			
			

			Cell cell1 = newrow.createCell(0);
			cell1.setCellValue(id);
			
			Cell cell2 = newrow.createCell(1);
			cell2.setCellValue(name);
			
			Cell cell3 = newrow.createCell(2);
			cell3.setCellValue(department);
			
			//write to file
			FileOutputStream out = new FileOutputStream(new File(wbook + ".xlsx"));
			workbook.write(out);
			System.out.println("new row added");
			out.close();
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
	
	public static void readexcel() throws IOException {
		
		try {
			FileInputStream file = new FileInputStream("employees.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet("records");
			
			//loop over rows in sheet
			Iterator<Row> rowiterator = sheet.rowIterator();
			while(rowiterator.hasNext()) {
				
				Row row = rowiterator.next();
				
				//loop over columns in each row
				Iterator<Cell> celliterator = row.cellIterator();
				while(celliterator.hasNext()) {
					Cell cell = celliterator.next();
					System.out.println(cell.getStringCellValue() + "\t");
				} //end column loop
				 System.out.println("\n");
			} // end row loop
			 file.close();
			System.out.println("------ end ------");

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	

	
     	public static void createworkbook() {
     		
     		//write to xlxs
			//create new instance workbook
			
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Employees");
			
			//Data
			Map<String, Object[]> data = new TreeMap<String, Object[]>();
			data.put("1", new Object[]{"id","name","department"});
			data.put("2", new Object[]{"1","joseph","mis"});
			data.put("3", new Object[]{"2","ryan","hr"});
			data.put("4", new Object[]{"3","juan","itds"});
		
	        Set<String> keyset = data.keySet();
	        
	        int rownum = 0;
	        
	        //loop each keyset
	        for(String key:keyset) {
	        	
	        	Row row = sheet.createRow(rownum+=1);
	        	Object[] obj = data.get(key);
	            int cellnum = 0;
	            //loop each column in each row

	            for(Object o:obj) {
            	Cell cell = row.createCell(cellnum++);
            	cell.setCellValue(o.toString());
            }
	            
	            //write file in filesystem
	           
	        }
	        try {
				
	        	FileOutputStream out = new FileOutputStream(new File("employees.xlsx"));
	        	workbook.write(out);
	        	out.close();
	        	System.out.println("write xlsx ok");

			} catch (Exception e) {
				e.printStackTrace();
			}
     		
     	}
	
		public static void createworkbook(String workbookname,String worksheetname){

			
			//write to xlxs
			//create new instance workbook
			
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Employees");
			
			//Data
			Map<String, Object[]> data = new TreeMap<String, Object[]>();
			data.put("1", new Object[]{"id","name","department"});
			data.put("2", new Object[]{"1","joseph","mis"});
			data.put("3", new Object[]{"2","ryan","hr"});
			data.put("4", new Object[]{"3","juan","itds"});
		
	        Set<String> keyset = data.keySet();
	        
	        int rownum = 0;
	        
	        //loop each keyset
	        for(String key:keyset) {
	        	
	        	Row row = sheet.createRow(rownum+=1);
	        	Object[] obj = data.get(key);
	            int cellnum = 0;
	            //loop each column in each row

	            for(Object o:obj) {
            	Cell cell = row.createCell(cellnum++);
            	cell.setCellValue(o.toString());
            }
	            
	            //write file in filesystem
	           
	        }
	        try {
				
	        	FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
	        	workbook.write(out);
	        	out.close();
	        	System.out.println("write xlsx ok");

			} catch (Exception e) {
				e.printStackTrace();
			}
			
			
		}
		

        
	}


