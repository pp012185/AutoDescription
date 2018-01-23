import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import static org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC;

public class AutoDescription_Items implements ICustomAction{
    @Override
    public ActionResult doAction(IAgileSession session, INode iNode, IDataObject obj) {

        System.out.println("------ Start ------");

        String filepath = "C:\\ExcelFile\\test_2.xlsx";
        String a="";
        try {
            String result = "";


            IItem item = (IItem) obj;     // 直接抓那 row 物件
            String itemNumber = item.getName();
            System.out.println();
            System.out.println("Part Number: " + itemNumber);



            ITable pending_tb = obj.getTable(ItemConstants.TABLE_PENDINGCHANGES);
            Iterator it = pending_tb.iterator();
            String status = "";
            String changeName = "";
            if (it.hasNext())
            {
                IRow row = (IRow) it.next();
                status = row.getValue(ItemConstants.ATT_PENDING_CHANGES_STATUS).toString();
                changeName = row.getValue(ItemConstants.ATT_PENDING_CHANGES_NUMBER).toString();
                System.out.println("status: "+status);
                System.out.println("change number: "+ changeName);
            }
            System.out.println("CE? "+(status.equals("CE")));
            System.out.println("PNR or PAR?" + (changeName.startsWith("PNR")||changeName.startsWith("PAR")));




            if(status.equals("CE") && (changeName.startsWith("PNR")||changeName.startsWith("PAR"))) {
                System.out.println("開始Auto Description");
                result+= readExcel(filepath, item);
                System.out.println("Result: " + result);
                if(!item.getValue(ItemConstants.ATT_TITLE_BLOCK_DESCRIPTION).toString().equals(result)) {
                    item.setValue(ItemConstants.ATT_TITLE_BLOCK_DESCRIPTION,result);
                }

                // a=readExcelList("10","EMB (Y/N)","No",filepath);
            }


        } catch (APIException e1) {
            e1.printStackTrace();
        } catch (Exception e1) {
            e1.printStackTrace();
        }
        System.out.println("------ End ------");
        return new ActionResult(0,"Success: ");
    }


    public static String readExcel(String path, IItem item)throws Exception {
        String result = "";
        String Description = "";
        try {
            // 用API Name去比對
            IAgileClass classes =item.getAgileClass();
            String subClass = classes.getAPIName();

            FileInputStream inp = new FileInputStream(path);
            XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
            XSSFSheet sheet = wb.getSheetAt(0);             //讀取wb內的頁面
            XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
            int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
            int columnlength = row.getPhysicalNumberOfCells();     // number of column

            for(int i = 1; i <rowlength; i++) {
                //System.out.println("開始找Subclass");
                String sub = "";   // Subclass
                String subNum = "";// Subclass number
                try {
                    row = sheet.getRow(i);
                    sub = row.getCell(2).toString();
                    subNum = sub.substring(1,3);
                    sub = sub.substring(1,3) + sub.substring(4,sub.length());
                    //DES = row.getCell(3).toString();
                }catch (NullPointerException e)
                {
                    e.getMessage();
                    e.printStackTrace();
                    continue;
                }
                subNum = subNum + ".0";
                //System.out.println("subnum= "+subNum);
                //System.out.println("Cell 3: "+sub );
                System.out.println(subClass.equals(sub)); // T or F
                if (subClass.equals(sub)) {                             // match subclass
                    for (int j = 4; j < columnlength; j++) {            // 看每個組成
                        String excelCell = row.getCell(j) + "";
                        if (!"null".equals(excelCell) && ("" + row.getCell(j)).length() != 0) { // excel field not null
                            if (excelCell.equals("end")) {              // end -> break
                                break;
                            } else if (excelCell.contains("$")) {       // e.g. $abc -> abc
                                System.out.println("$: " + getString_DollarSigns(excelCell));
                                Description += getString_DollarSigns(excelCell);
                                continue;
                            } else if (excelCell.contains("*")) {
                                String preexcelCell = row.getCell(j - 1) + "";            // 抓前一個欄位名稱
                                String postexcelCell = row.getCell(j + 1) + "";       // 抓後一個欄位名稱
                                String value = getString_Asterisk(excelCell, item, preexcelCell, postexcelCell);
                                System.out.println("Asterisk: " + value);
                                Description += value;
                                System.out.println("**" + excelCell + "::Value:" + value + ",ProductName:" + Description);
                                continue;
                            }else if(excelCell.equals("Process Type") || excelCell.equals("Process Type!")){
                                String value = getProcessType(excelCell,item);
                                System.out.println("Process Type: " + value);
                                Description += value;
                                System.out.println("**" + excelCell+ "::Value:" +value + ",ProductName:" + Description);
                                continue;
                            }else if (item.getCell("Page Three." + excelCell) == null && (!excelCell.contains("!"))) { // System no this field  (要用getCell !)
                                System.out.println("no field:" + row.getCell(j));
                                Description += "█" + item + ":no field:" + row.getCell(j) + " ";
                                continue;
                            } else if(excelCell.contains("!") && item.getCell("Page Three." + excelCell.substring(0,excelCell.length()-1)) == null){
                                System.out.println("no field:" + row.getCell(j));
                                Description += "█" + item + ":no field:" + row.getCell(j) + " ";
                                continue;
                            } else {// excel field==PLM field
                                String value = getString_NoSigns(subNum,excelCell,item, path);
                                System.out.println("No Signs: " + value);
                                Description += value;
                                System.out.println("**" + excelCell+ "::Value:" +value + ",ProductName:" + Description);
                                continue;
                            }
                        }
                    } // end Column
                } // end match subclass
                if (!Description.equals("")) {
                    result = Description;
                    //System.out.println(result);
                    //System.out.println(DES);
                    break;
                }
            }
        }catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    private static String getString_DollarSigns (String excelCell) throws Exception
    {
        String value ="";
        if(excelCell.contains("!")){
            value = excelCell.substring(1, (excelCell.length()-1));
        }else{
            value = excelCell.substring(1, excelCell.length()) + " ";
        }
        return value;
    }

    private static String getString_NoSigns (String subNum, String excelCell, IItem item, String path) throws Exception
    {
        String value = "";
        if(excelCell.contains("!")){
            String NewExcelCell = excelCell.substring(0, excelCell.length() - 1);
            IAgileClass cls = item.getAgileClass();
            IAttribute atr = cls.getAttribute("Page Three." + NewExcelCell);
            System.out.println(atr.getDataType());
            if (atr.getDataType() == 2 || atr.getDataType() == 8) {                                                   // 組成為text or numeric
                if (item.getValue("Page Three." + NewExcelCell) == ""  || item.getValue("Page Three." + NewExcelCell) == null) {     // 若有對應的text，但是沒值
                    System.out.println("Field Name:" + NewExcelCell + " -> No Value");
                } else {
                    String ProductNameValue = item.getValue("Page Three." + NewExcelCell) + "";
                    value += ProductNameValue ;
                }
            } else if (atr.getDataType() == 4) {                                            // 組成為list，先找Excel,沒有找系統，還是沒有回覆""
                if (item.getValue("Page Three." + NewExcelCell) == "" || item.getValue("Page Three." + NewExcelCell) == null) {     // 若有對應的list，但是沒值
                    System.out.println("Field Name:" + NewExcelCell + " -> No Value");
                }else{
                    // String tmp2 = item.getValue("Page Three." + NewExcelCell).toString();
                    // ICell listCell2 = item.getCell("Page Three." + NewExcelCell);
                    // IAgileList list2 = (IAgileList) listCell2.getValue();
                    // String ProductNameValue = ((IAgileList) list2.getChild(tmp2)).getDescription();               // get the description of option in the list
                    String ProductNameValue = readExcelList(subNum , NewExcelCell, item.getValue("Page Three." + NewExcelCell).toString(), path);
                    System.out.println("List item Description Value: "+ProductNameValue);
                    if (ProductNameValue == "0") {                                         // excel裡沒有對應欄位，抓系統值
                        ProductNameValue = item.getValue("Page Three." + NewExcelCell) + "";
                        value += ProductNameValue;
                        System.out.println("Field Name:" + NewExcelCell );
                    }else {
                        value += ProductNameValue;
                    }
                }
            }
        }else{
            IAgileClass cls = item.getAgileClass();
            IAttribute atr = cls.getAttribute("Page Three." + excelCell);
            System.out.println(atr.getDataType());
            if (atr.getDataType() == 2 || atr.getDataType() == 8) {                                              // 組成為text or numeric
                if (item.getValue("Page Three." + excelCell) == "" || item.getValue("Page Three." + excelCell) == null) {   // 若有對應的text，但是沒值
                    System.out.println("Field Name:" + excelCell + " -> No Value");
                } else {
                    String ProductNameValue = item.getValue("Page Three." + excelCell) + "";
                    value += ProductNameValue + " ";
                }
            } else if (atr.getDataType() == 4) {                                        // 組成為list，list加其description 不直接加 name
                if (item.getValue("Page Three." + excelCell) == "" || item.getValue("Page Three." + excelCell) == null) {    // 若有對應的list，但是沒值
                    System.out.println("Field Name:" + excelCell + " -> No Value");
                }else{
                    // String tmp2 = item.getValue("Page Three." + excelCell).toString();
                    // ICell listCell2 = item.getCell("Page Three." + excelCell);
                    // IAgileList list2 = (IAgileList) listCell2.getValue();
                    // String ProductNameValue = ((IAgileList) list2.getChild(tmp2)).getDescription();          // get the description of option in the list
                    String ProductNameValue = readExcelList(subNum , excelCell, item.getValue("Page Three." + excelCell).toString(), path);
                    System.out.println("List item Description Value: "+ProductNameValue);
                    if (ProductNameValue == "0") {                                     // excel裡沒有對應欄位，抓系統值
                        ProductNameValue = item.getValue("Page Three." + excelCell) + "";
                        value += ProductNameValue+ " ";
                        System.out.println("Field Name:" + excelCell );
                    }else if(ProductNameValue==""){
                        value += ProductNameValue;
                    } else{
                        value += ProductNameValue + " ";
                    }
                }
            }
        }
        return value;
    }

    private static String getString_Asterisk (String excelCell, IItem item, String preexcelCell, String postexcelCell) throws Exception
    {
        String value = "";
        if(excelCell.contains("!")){
            //System.out.println(excelCell+" Contain *,! ");
            //System.out.println("頭: "+excelCell.substring(0,1));
            //System.out.println("尾: "+excelCell.substring(excelCell.length()-1,excelCell.length()));
            //System.out.println("尾2: "+excelCell.substring(excelCell.length()-2,excelCell.length()-1));
            if (excelCell.substring(0,1).equals("*")){                                                // "*"號在第一個 => 看前一個組欄位
                //System.out.println("進入 頭為*");
                if(preexcelCell.contains("!")) preexcelCell = preexcelCell.substring(0,preexcelCell.length()-1);
                //System.out.println("pre excel cell: "+preexcelCell);
                if(item.getCell("Page Three."+preexcelCell)==null){                               // 前一欄在系統不存在
                    //System.out.println("前一欄在系統中不存在");
                }else if(item.getValue("Page Three." +preexcelCell)==null || item.getValue("Page Three." +preexcelCell)==""){                         // 前一欄位沒有值
                    //System.out.println("前一欄沒值");
                }else {                                                                               // 前一個欄位有值 => 當字串填入
                    //System.out.println("前一欄值: "+item.getValue("Page Three." +preexcelCell));
                    value = excelCell.substring(1, excelCell.length()-1);
                }
            }else if(excelCell.substring(excelCell.length()-2,excelCell.length()-1).equals("*")) {    // "*"號在倒數第二個 => 看後一個欄位
                //System.out.println("進入 尾為*");
                if (postexcelCell.contains("!"))
                    postexcelCell = postexcelCell.substring(0, postexcelCell.length() - 1);
                //System.out.println("post excel cell: " + postexcelCell);
                if(item.getCell("Page Three."+postexcelCell)==null){                              // 後一欄在系統不存在
                    //System.out.println("後一欄在系統中不存在");
                }else if(item.getValue("Page Three." +postexcelCell)==null || item.getValue("Page Three." +postexcelCell)==""){                        // 後一欄位沒有值
                    //System.out.println("後一欄沒值");
                }else {                                                                               // 後一個欄位有值 => 當字串填入
                    //System.out.println("後一欄值: "+item.getValue("Page Three." +postexcelCell));
                    value = excelCell.substring(0, (excelCell.length()-2));
                }
            }
        }else{
            //System.out.println(excelCell+" Contain *");
            //System.out.println("頭: "+excelCell.substring(0,1));
            //System.out.println("尾: "+excelCell.substring(excelCell.length()-1,excelCell.length()));
            if (excelCell.substring(0,1).equals("*")){                                                // "*"號在第一個 => 看前一個組欄位
                //System.out.println("進入 頭為*");
                if(preexcelCell.contains("!")) preexcelCell = preexcelCell.substring(0,preexcelCell.length()-1);
                //System.out.println("pre excel cell: "+preexcelCell);
                if(item.getCell("Page Three."+preexcelCell)==null){                               // 前一欄在系統找不到
                    //System.out.println("前一欄在系統中找不到!!");
                }else if(item.getValue("Page Three." +preexcelCell)==null || item.getValue("Page Three." +preexcelCell)==""){                         // 前一欄位沒有值
                    //System.out.println("前一欄沒值!!");
                }else {                                                                               // 前一個欄位有值 => 當字串填入
                    //System.out.println("前一欄值: "+item.getValue("Page Three." +preexcelCell));
                    value = excelCell.substring(1, excelCell.length()) + " ";
                }
            }else if(excelCell.substring(excelCell.length()-1,excelCell.length()).equals("*")){       // "*"號在最後一個 => 看後一個欄位
                //System.out.println("進入 尾為*");
                if(postexcelCell.contains("!")) postexcelCell = postexcelCell.substring(0,postexcelCell.length()-1);
                //System.out.println("post excel cell: "+postexcelCell);
                if(item.getCell("Page Three."+postexcelCell)==null){                              // 後一欄在系統找不到
                    //System.out.println("後一欄在系統中找不到!!");
                }else if(item.getValue("Page Three." +postexcelCell)==null || item.getValue("Page Three." +postexcelCell)==""){                        // 後一欄位沒有值
                    //System.out.println("後一欄沒值!!");
                }else {                                                                               // 後一個欄位有值 => 當字串填入
                    //System.out.println("後一欄值: "+item.getValue("Page Three." +postexcelCell));
                    value = excelCell.substring(0, (excelCell.length()-1)) + " ";
                }
            }
        }
        return value;
    }

    private static String SplitString(String ProductNameValue)
    {
        String[] value =  ProductNameValue.split("｜");
        return value[1];
    }

    private static String readExcelList(String subclassNum, String fieldName, String fieldNameValue,String path) throws IOException {
        String value="";
        int inExcel = 0; // 0 -> no, 1 -> yes
        FileInputStream inp = new FileInputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook(inp);                //讀取Excel
        XSSFSheet sheet = wb.getSheetAt(1);             //讀取wb內的頁面
        XSSFRow row = sheet.getRow(0);               //讀取頁面0的第一行
        int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
        int columnlength = row.getPhysicalNumberOfCells();     // number of column
        //System.out.println("rowlength"+rowlength);
        //System.out.println("columnlength"+columnlength);
        int count1 = 0;
        int count2 = 0;


        for(int i=1;i< rowlength;i++) {
            row = sheet.getRow(i);
            Cell cell = row.getCell(0);           // excel cell is numeric => 10.0
            //cell.setCellType(Cell.CELL_TYPE_STRING);    // change to string => 10
            if (cell.toString().equals(subclassNum) ) count1++;
        }
        System.out.println("count1: "+ count1);
        for(int i=1;i< rowlength;i++){
            row = sheet.getRow(i);
            Cell cell = row.getCell(0);
            //cell.setCellType(Cell.CELL_TYPE_STRING);
            if (cell.toString().equals(subclassNum)) {
                //System.out.println("i = "+i);
                for (int j = i; j < i+count1; j++) {
                    row = sheet.getRow(j);
                    if (row.getCell(1).toString().equals(fieldName)) count2++;
                }
                System.out.println("count2: "+ count2);
                for (int j = i; j < i+count1; j++) {
                    row = sheet.getRow(j);
                    if (row.getCell(1).toString().equals(fieldName))
                    {
                        //System.out.println("j = "+j);
                        for(int k = j; k < j+count2; k++)
                        {
                            row = sheet.getRow(k);
                            System.out.println(row.getCell(2).toString());
                            if(row.getCell(2).toString().equals(fieldNameValue))
                            {
                                System.out.println(row.getCell(3));
                                if (row.getCell(3) == null) {
                                    value = "";
                                    inExcel =1;
                                }
                                else{
                                    value = row.getCell(3).toString();
                                    inExcel=1;
                                }
                                //System.out.println("k ="+ k + ": " + value);
                                break;
                            }
                        }
                        break;
                    }
                }
                break;
            }
        }
        if (value=="" && inExcel==0) value="0";
        //System.out.println("value:"+value+" inexcle:"+inExcel);
        return value;
    }

    private static String removespace(String col_name) {
        String value ="";
        value = col_name.replaceAll("\\s+","");
        return value;
    }

    private static String getProcessType(String excelCell,IItem item) throws APIException {
        String value="";
        if(excelCell.contains("!")) {
            String NewExcelCell = excelCell.substring(0, excelCell.length() - 1);
            if (item.getValue("Page Two." + NewExcelCell) == "" || item.getValue("Page Two." + NewExcelCell) == null) {     // 若有對應的list，但是沒值
                System.out.println("Field Name:" + NewExcelCell + " -> No Value");
            } else {
                String ProductNameValue = item.getValue("Page Two." + NewExcelCell) + "";
                System.out.println("List item Description Value: " + ProductNameValue);
                value += ProductNameValue;
            }
        }else {
            if (item.getValue("Page Two." + excelCell) == "" || item.getValue("Page Two." + excelCell) == null) {     // 若有對應的list，但是沒值
                System.out.println("Field Name:" + excelCell + " -> No Value");
            } else {
                String ProductNameValue = item.getValue("Page Two." + excelCell) + "";
                System.out.println("List item Description Value: " + ProductNameValue);
                value += ProductNameValue;
            }
        }
        if(value.equals("SMT")) value = "SMD";
        return value;
    }

}
