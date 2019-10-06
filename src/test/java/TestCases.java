import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.Test;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public class TestCases extends BaseTestCases {
    ConcurrentHashMap<Integer,String> resultMap=new ConcurrentHashMap<Integer, String>();
    @Test(dataProvider="excelDataProvider",dataProviderClass = GenericExcelDataProvider.class)
    public void test1(int testCaseId, String testName)
    {

        resultMap.put(testCaseId,"Failed");

    }

    @AfterSuite
   public  void AfterSuite() throws IOException
    {
        FileInputStream fileInp=new FileInputStream("testData/testCaseExcel.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(fileInp);
        XSSFSheet sheet=workbook.getSheet("TestData1");
        XSSFRow initialRow=sheet.getRow(0);
        int lastColumnNumber=initialRow.getLastCellNum();
        XSSFCell cell;
        int testCaseId;
        String result;


        //For getting current date to save with results
        SimpleDateFormat formater=new SimpleDateFormat("YYYY/MM/DD HH:mm");
        Date d=new Date();
        String currentDate=formater.format(d);
        //Writing Column header in 0th row by creating a cell at last

        cell=initialRow.createCell(lastColumnNumber+1);
        cell.setCellValue("Result"+ currentDate);

        //Iterating throughout the resultMap to set the values of result in the sheet
        Iterator iterator=resultMap.entrySet().iterator();
        while(iterator.hasNext())
        {
            Map.Entry mapEntry= (Map.Entry)iterator.next();
            testCaseId=(Integer)mapEntry.getKey();
            result=(String)mapEntry.getValue();

            //Updating in excelCell
            cell=sheet.getRow(testCaseId).createCell(lastColumnNumber+1);
            cell.setCellValue(result);
        }
        //Writing back to excel
        FileOutputStream fout=new FileOutputStream(new File("testData/testCaseExcel.xlsx"));
        workbook.write(fout);
        fileInp.close();
        fout.close();

    }


}

