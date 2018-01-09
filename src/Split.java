import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.FileInputStream;
import java.io.IOException;

public class Split {
    public static void main(String[] args) throws IOException {
        /*String str = "5780｜QC";
        System.out.println(str.contains("If"));
        String[] a = str.split("｜");
        System.out.println(a);

        System.out.println(SplitString(str));
        String c = "*V";
        System.out.println("TEST "+c.substring(0,1));
        System.out.println(c.substring(0,1).equals("*"));
        System.out.println((c.contains("*")));
        */
        String path = "C:\\Users\\pp\\Desktop\\Quixant\\test_list.xlsx";
        String result = testreadExcelList("A",path);
        System.out.println("readExcelList: "+result);
    }



    private static String SplitString(String ProductNameValue)
    {
        String[] value =  ProductNameValue.split("｜");
        return value[1];
    }

    private static String testreadExcelList(String SubclassName, String path) throws IOException {
        String value="";

        FileInputStream inp = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
        XSSFSheet sheet = wb.getSheetAt(1);             //讀取wb內的頁面
        XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
        int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
        int columnlength = row.getPhysicalNumberOfCells();     // number of column
        System.out.println("有 "+rowlength+" 列");
        System.out.println("有 "+columnlength+" 行");

        return value;
    }
}
