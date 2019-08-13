import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public ArrayList<String> getData(String testCaseName) throws IOException
	{		
		//defining the arraylist
		
		ArrayList<String> aList = new ArrayList<String>();
		//file input stream argument
		FileInputStream fis = new FileInputStream("C:\\Users\\pratikd\\Desktop\\ExcelFileData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();
		for(int i=0; i<sheets; i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				//identify testcases column by scanning the entire 1st row
				
				Iterator<Row> rows = sheet.iterator(); //sheet is a collection of rows
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator(); // row is a collection of cells
				
				int k=0;
				int column=0;
				
				while(ce.hasNext())
				{
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase(testCaseName)) 
					{
						column = k;
						System.out.println("hi");
					}
					
					k++;
				}
				System.out.println(column);
				
				//once column is identified then scan entire testcase column to identify the purchase testcase row
				while(rows.hasNext())
				{
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) 
					{
						//after you grab purchase testcase row pull all the data of that row and feed into test
						Iterator<Cell> cv= r.cellIterator();
						while(cv.hasNext())
						{
							Cell c = cv.next();
							if(c.getCellType() == CellType.STRING)
							{
								aList.add(c.getStringCellValue());
							}
							else
							{
								aList.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
				}
			}
		}
		return aList;
	}
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	}

}
