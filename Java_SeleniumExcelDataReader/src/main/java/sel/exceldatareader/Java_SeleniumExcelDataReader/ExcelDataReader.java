package sel.exceldatareader.Java_SeleniumExcelDataReader;

/**
 * Hello world!
 *
 */
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataReader{

	public static  FileInputStream fis;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static int rowCount;
	
	public static void GetExcelWorkbook(String ExcelName) 
	{
		try 
		{
			fis = new FileInputStream("./data/"+ExcelName+".xlsx");
			System.out.println("ExcelWorkbook is Identified successfully");
		}
		catch(Exception ex)
		{
			System.out.println("ExcelWorkbook not identified : Error - "+ex.getMessage());
		}
	}

	public static void GetSheet(String sheetName) 
	{
		try 
		{
			workbook = new XSSFWorkbook(fis);
			 sheet = workbook.getSheet(sheetName);
			//XSSFSheet sheet = workbook.getSheetAt(0);
			System.out.println("ExcelSheet is Identified successfully");
		}
		catch(Exception ex)
		{
			System.out.println("ExcelSheet not identified : Error - "+ex.getMessage());
		}
	}
	
	public static String[][] ReadData(String ExcelName, String sheetName) 
	{
		String[][] data = null ;
		GetExcelWorkbook(ExcelName);
		GetSheet(sheetName);
		
		try 
		{				
			// get the number of rows
			rowCount = sheet.getLastRowNum();

			// get the number of columns
			int columnCount = sheet.getRow(0).getLastCellNum();
			data = new String[rowCount][columnCount];

			System.out.println(rowCount);
			// loop through the rows
			for(int i=1; i <rowCount+1; i++){
				try 
				{
					row = sheet.getRow(i);
					for(int j=0; j <columnCount; j++)
					{
						// loop through the columns
						try 
						{
							String cellValue = "";
							try
							{
								//	cellValue = row.getCell(j).getStringCellValue();
								CellType cellType = row.getCell(j).getCellTypeEnum();	
								if (cellType != CellType.STRING) 
								{
									row.getCell(j).setCellType(CellType.STRING);
								}
								cellValue = row.getCell(j).getStringCellValue();
							}
							catch(NullPointerException ex)
							{
								System.out.println("Data Reading : Error No Data exists in the cell");
							}

							data[i-1][j]  = cellValue; // add to the data array
						} 
						catch (Exception ex) 
						{
							System.out.println("Data Reading : Error in 'For Loop' Looping through Columns");
							ex.printStackTrace();
						}				
					}
				} 
				catch (Exception ex)
				{
					System.out.println("Data Reading : Error in 'For Loop' Looping through Rows- "+ex.getMessage());
					ex.printStackTrace();
				}
			}
			fis.close();
			workbook.close();
		} 
		catch (Exception ex) {
			System.out.println("No of Rows Reading : Error in fething No of Rows- "+ex.getMessage());
			ex.printStackTrace();
		}

		return data;

	}

}
