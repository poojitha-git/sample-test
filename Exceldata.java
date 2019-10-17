package Generic_Library;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exceldata {
	
	FileInputStream fs;
	XSSFWorkbook wb;
	
//	Initializing the Excel
	
	public Exceldata(String fpath) throws Exception{
		
		fs = new FileInputStream(fpath);
		wb = new XSSFWorkbook(fs);
	}
	
//	Get the Row count
	
	public int getrowcount(String sheetname){
		
		XSSFSheet x = wb.getSheet(sheetname);
		return x.getLastRowNum();
	}
	
//	Get the Column count
	
	public int getcolumncount(String sheetname){
		
		XSSFSheet x = wb.getSheet(sheetname);
		return x.getRow(0).getLastCellNum();
	}	
//	Read cell values
	
	@SuppressWarnings("deprecation")
	public String readvalue(int ri,int ci,String sheetname){
				
		XSSFSheet x = wb.getSheet(sheetname);
		
		XSSFCell xc = x.getRow(ri).getCell(ci);
				
		String cellvalue = null;
		
		if(xc.getCellType() == xc.CELL_TYPE_STRING){
			
			cellvalue = xc.getStringCellValue();
//			System.out.println("Celvalue: "+cellvalue);
		
		}else if(xc.getCellType() == xc.CELL_TYPE_NUMERIC){
		
			cellvalue = String.valueOf(xc.getNumericCellValue());
			
		}else if(xc.getCellType() == xc.CELL_TYPE_BLANK && xc.getCellType()==xc.CELL_TYPE_ERROR){
			
			cellvalue = "";
		}
		return cellvalue;
	}
	
//	Write cell values
	
	public void writevalue(int ri,int ci,String sheetname,String input) throws Exception{
		
		XSSFSheet x = wb.getSheet(sheetname);
		//System.out.println("ExcelRe"+input);
		x.getRow(ri).getCell(ci).setCellValue(input);

		
	}
	
	//Save and close the streams
	
	public void saveclose(String fpath) throws Exception{
		
		FileOutputStream fo = new FileOutputStream(fpath);
	
//	Write excel
		wb.write(fo);
		
//		Close stream
		
		fo.close();
		fs.close();
	}



}
