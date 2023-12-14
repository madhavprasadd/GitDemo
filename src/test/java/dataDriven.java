import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	// first create fileinpitstream for the file
	// get access to the sheet we required

	// once clm is identified then scan entire clm to get the desired data

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
        ArrayList<String> a =new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C://Users//Gauranga//OneDrive//Documents//sel//Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheet = workbook.getNumberOfSheets();
		for (int i = 0; i < sheet; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) 
			{
				XSSFSheet sheet1 = workbook.getSheetAt(i);
				
				// scan entire firstrow to identify the desired column
				
			Iterator <Row>	rows=sheet1.iterator();
			Row firstrow=rows.next();
			Iterator<Cell> ce=firstrow.cellIterator();
			int k=0;
			int column=0;
			while(ce.hasNext())
			{
				Cell value=ce.next();
				if(value.getStringCellValue().equalsIgnoreCase("Testcase"))
						{
										column=k;
										
						}
				k++;
			}
			      while(rows.hasNext())
			      {
			    	  Row r=rows.next();
			    	  if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase"))
			    	  {
			    		 Iterator<Cell> cv= r.cellIterator();
			    		 while(cv.hasNext())
			    		 {
			    			a.add(cv.next().getStringCellValue());
			    			
			    		 }
			    	  }

			    	 
			      }
			}

		}

	}

}
