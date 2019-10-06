import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class GenericExcelDataProvider {
    @DataProvider(name="excelDataProvider")
    public Object[][] ReadTestDataFromExcel() throws FileNotFoundException, IOException
    {
        FileInputStream fileInp = new FileInputStream("testData/testCaseExcel.xlsx");
        XSSFWorkbook workbook= new XSSFWorkbook(fileInp);
        XSSFSheet sheet = workbook.getSheet("TestData1");
        int noOfRows=sheet.getPhysicalNumberOfRows();
        int noOfColumn=2;
        Object[][] testData=new Object[noOfRows-1][noOfColumn];
        for(int i=0;i<noOfRows-1; i++)
        {
            XSSFRow row=sheet.getRow(i+1);
            for (int j=0; j<noOfColumn;j++)
            {
                XSSFCell cell=row.getCell(j);
                if(cell.getCellType()== CellType.NUMERIC)
                {
                    testData[i][j]=(int)(cell.getNumericCellValue());
                }
                else if(cell.getCellType()==CellType.STRING)
                {
                    testData[i][j]=cell.getStringCellValue();
                }
            }
        }
        fileInp.close();
        return testData;
    }
}

