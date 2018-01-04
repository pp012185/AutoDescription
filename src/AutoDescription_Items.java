import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAction;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class AutoDescription_Items implements ICustomAction{
    @Override
    public ActionResult doAction(IAgileSession session, INode iNode, IDataObject obj) {

        System.out.println("------ Start ------");

        String filepath = "C:\\ExcelFile\\test_2.xlsx";
        try {

                String result = "";

                IItem item = (IItem) obj;     // 直接抓那 row 物件
                String itemNumber = item.getName();
                System.out.println();
                System.out.println("Part Number: " + itemNumber);

                result+= readExcel(filepath, item);
                item.setValue(ItemConstants.ATT_TITLE_BLOCK_DESCRIPTION,result);

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
            XSSFWorkbook wb = new XSSFWorkbook(inp);       //讀取Excel
            XSSFSheet sheet = wb.getSheetAt(0);    //讀取wb內的頁面
            XSSFRow row = sheet.getRow(0);      //讀取頁面0的第一行
            int rowlength = sheet.getPhysicalNumberOfRows();       // number of row
            int columnlength = row.getPhysicalNumberOfCells();     // number of column

            for(int i = 1; i <rowlength; i++) {
                System.out.println("開始找Subclass");
                String sub = "";
                String DES = "";
                try {
                    row = sheet.getRow(i);
                    sub = row.getCell(2).toString();
                    sub = sub.substring(1,3) + sub.substring(4,sub.length());
                    DES = row.getCell(3).toString();
                }catch (NullPointerException e)
                {
                    e.getMessage();
                    e.printStackTrace();
                    continue;
                }
                System.out.println("DES: "+DES);
                System.out.println("Cell 3: "+sub );
                System.out.println(subClass.equals(sub)); // T or F

                if (subClass.equals(sub)) { // match subclass

                    for (int j = 4; j < columnlength; j++) {          // 看每個組成
                        String excelCell = row.getCell(j) + "";
                        if (!"null".equals(excelCell) && ("" + row.getCell(j)).length() != 0) { // excel field not null
                            if (excelCell.equals("end")) {
                                break;
                            } else if (excelCell.contains("$")) {            // e.g. $abc -> abc
                                if(excelCell.contains("!")){
                                    String value = excelCell.substring(1, (excelCell.length()-1));
                                    Description += value ;
                                    continue;
                                }else{
                                    String value = excelCell.substring(1, excelCell.length());
                                    Description += value + " ";
                                    continue;
                                }
                            }else if(excelCell.contains("*")){
                                if(excelCell.contains("!")){
                                    System.out.println(excelCell+" Contain *,! ");
                                    System.out.println("頭: "+excelCell.substring(0,1));
                                    System.out.println("尾: "+excelCell.substring(excelCell.length()-1,excelCell.length()));
                                    System.out.println("尾2: "+excelCell.substring(excelCell.length()-2,excelCell.length()-1));
                                    if (excelCell.substring(0,1).equals("*")){                                                // "*"號在第一個 => 看前一個組欄位
                                        System.out.println("進入 頭為*");
                                        String preexcelCell = row.getCell(j-1)+"";                                  // 抓前一個欄位名稱
                                        if(preexcelCell.contains("!")) preexcelCell = preexcelCell.substring(0,preexcelCell.length()-1);
                                        System.out.println("pre excel cell: "+preexcelCell);
                                        if(item.getCell("Page Three."+preexcelCell)==null){                               // 前一欄在系統找不到
                                            System.out.println("前一欄在系統中找不到!!");
                                            continue;
                                        }else if(item.getValue("Page Three." +preexcelCell)==""){                         // 前一欄位沒有值
                                            System.out.println("前一欄沒值!!");
                                            continue;
                                        }else {                                                                               // 前一個欄位有值 => 當字串填入
                                            System.out.println("前一欄值: "+item.getValue("Page Three." +preexcelCell));
                                            String value = excelCell.substring(1, excelCell.length()-1);
                                            Description += value;
                                            continue;
                                        }
                                    }else if(excelCell.substring(excelCell.length()-2,excelCell.length()-1).equals("*")) {    // "*"號在倒數第二個 => 看後一個欄位
                                        System.out.println("進入 尾為*");
                                        String postexcelCell = row.getCell(j + 1) + "";                             // 抓後一個欄位名稱
                                        if (postexcelCell.contains("!"))
                                            postexcelCell = postexcelCell.substring(0, postexcelCell.length() - 1);
                                        System.out.println("post excel cell: " + postexcelCell);
                                        if(item.getCell("Page Three."+postexcelCell)==null){                              // 後一欄在系統找不到
                                            System.out.println("後一欄在系統中找不到!!");
                                            continue;
                                        }else if(item.getValue("Page Three." +postexcelCell)==""){                        // 後一欄位沒有值
                                            System.out.println("後一欄沒值!!");
                                            continue;
                                        }else {                                                                               // 後一個欄位有值 => 當字串填入
                                            System.out.println("後一欄值: "+item.getValue("Page Three." +postexcelCell));
                                            String value = excelCell.substring(0, (excelCell.length()-2));
                                            Description += value ;
                                            continue;
                                        }
                                    }
                                }else{
                                    System.out.println(excelCell+" Contain *");
                                    System.out.println("頭: "+excelCell.substring(0,1));
                                    System.out.println("尾: "+excelCell.substring(excelCell.length()-1,excelCell.length()));
                                    if (excelCell.substring(0,1).equals("*")){                                                // "*"號在第一個 => 看前一個組欄位
                                        System.out.println("進入 頭為*");
                                        String preexcelCell = row.getCell(j-1)+"";                                  // 抓前一個欄位名稱
                                        if(preexcelCell.contains("!")) preexcelCell = preexcelCell.substring(0,preexcelCell.length()-1);
                                        System.out.println("pre excel cell: "+preexcelCell);
                                        if(item.getCell("Page Three."+preexcelCell)==null){                               // 前一欄在系統找不到
                                            System.out.println("前一欄在系統中找不到!!");
                                            continue;
                                        }else if(item.getValue("Page Three." +preexcelCell)==""){                         // 前一欄位沒有值
                                            System.out.println("前一欄沒值!!");
                                            continue;
                                        }else {                                                                               // 前一個欄位有值 => 當字串填入
                                            System.out.println("前一欄值: "+item.getValue("Page Three." +preexcelCell));
                                            String value = excelCell.substring(1, excelCell.length());
                                            Description += value + " ";
                                            continue;
                                        }
                                    }else if(excelCell.substring(excelCell.length()-1,excelCell.length()).equals("*")){       // "*"號在最後一個 => 看後一個欄位
                                        System.out.println("進入 尾為*");
                                        String postexcelCell = row.getCell(j+1)+"";                                 // 抓後一個欄位名稱
                                        if(postexcelCell.contains("!")) postexcelCell = postexcelCell.substring(0,postexcelCell.length()-1);
                                        System.out.println("post excel cell: "+postexcelCell);
                                        if(item.getCell("Page Three."+postexcelCell)==null){                              // 後一欄在系統找不到
                                            System.out.println("後一欄在系統中找不到!!");
                                            continue;
                                        }else if(item.getValue("Page Three." +postexcelCell)==""){                        // 後一欄位沒有值
                                            System.out.println("後一欄沒值!!");
                                            continue;
                                        }else {                                                                               // 後一個欄位有值 => 當字串填入
                                            System.out.println("後一欄值: "+item.getValue("Page Three." +postexcelCell));
                                            String value = excelCell.substring(0, (excelCell.length()-1));
                                            Description += value + " ";
                                            continue;
                                        }
                                    }
                                }


                            } else if (item.getCell("Page Three." + excelCell) == null && (!excelCell.contains("!"))) { // System no this field  (要用getCell !)
                                //System.out.println(item.getCell(1541).getName());
                                System.out.println("no field:" + row.getCell(j));
                                Description += "█" + item + ":no field:" + row.getCell(j) + " ";
                                continue;
                            } else if(excelCell.contains("!") && item.getCell("Page Three." + excelCell.substring(0,excelCell.length()-1)) == null){
                                //System.out.println(item.getCell(1541).getName());
                                System.out.println("no field:" + row.getCell(j));
                                Description += "█" + item + ":no field:" + row.getCell(j) + " ";
                                continue;
                            } else {// excel field==PLM field
                                if(excelCell.contains("!")){
                                    String NewExcelCell = excelCell.substring(0, excelCell.length() - 1);
                                    IAgileClass cls = item.getAgileClass();
                                    IAttribute atr = cls.getAttribute("Page Three." + NewExcelCell);
                                    System.out.println(atr.getDataType());
                                    if (atr.getDataType() == 2) { // 組成為text
                                        if (item.getValue("Page Three." + NewExcelCell).toString() == "") {   // 若有對應的text，但是沒值
                                            System.out.println("Field Name:" + NewExcelCell + " -> No Value");
                                            continue;
                                        } else {
                                            String ProductNameValue = item.getValue("Page Three." + NewExcelCell) + "";
                                            Description += ProductNameValue ;
                                            System.out.println("**" + NewExcelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                            continue;
                                        }
                                    } else if (atr.getDataType() == 4) {                                            // 組成為list   // list 加其description 不直接加 name
                                        if (item.getValue("Page Three." + NewExcelCell).toString() == "") {     // 若有對應的list，但是沒值
                                            System.out.println("Field Name:" + NewExcelCell + " -> No Value");
                                            continue;
                                        }else{
                                            String tmp2 = item.getValue("Page Three." + NewExcelCell).toString();
                                            ICell listCell2 = item.getCell("Page Three." + NewExcelCell);
                                            IAgileList list2 = (IAgileList) listCell2.getValue();
                                            String ProductNameValue = ((IAgileList) list2.getChild(tmp2)).getDescription();               // get the description of option in the list
                                            System.out.println("List item Description Value: "+ProductNameValue);
                                            if (ProductNameValue == null) {             // list item 的description 沒有值
                                                ProductNameValue = "";
                                                Description += ProductNameValue;
                                                System.out.println("**" + NewExcelCell + "::List item Description Value:" + ProductNameValue + ",ProductName:" + Description);
                                                continue;
                                            }
                                            Description += ProductNameValue;
                                            System.out.println("**" + NewExcelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                            continue;
                                        }
                                    }

                                }else{
                                    IAgileClass cls = item.getAgileClass();
                                    IAttribute atr = cls.getAttribute("Page Three." + excelCell);
                                    System.out.println(atr.getDataType());
                                    if (atr.getDataType() == 2) { // 組成為text
                                        if (item.getValue("Page Three." + excelCell).toString() == "") {   // 若有對應的text，但是沒值
                                            System.out.println("Field Name:" + excelCell + " -> No Value");
                                            continue;
                                        } else {
                                            String ProductNameValue = item.getValue("Page Three." + excelCell) + "";
                                            Description += ProductNameValue + " ";
                                            System.out.println("**" + excelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                            continue;
                                        }
                                    } else if (atr.getDataType() == 4) {                                        // 組成為list   // list 加其description 不直接加 name
                                        if (item.getValue("Page Three." + excelCell).toString() == "") {    // 若有對應的list，但是沒值
                                            System.out.println("Field Name:" + excelCell + " -> No Value");
                                            continue;
                                        }else{
                                            String tmp2 = item.getValue("Page Three." + excelCell).toString();
                                            ICell listCell2 = item.getCell("Page Three." + excelCell);
                                            IAgileList list2 = (IAgileList) listCell2.getValue();
                                            String ProductNameValue = ((IAgileList) list2.getChild(tmp2)).getDescription();          // get the description of option in the list
                                            System.out.println("List item Description Value: "+ProductNameValue);
                                            if (ProductNameValue == null) {
                                                ProductNameValue = "";
                                                Description += ProductNameValue;
                                                System.out.println("**" + excelCell + "::List item Description Value: " + ProductNameValue + ",ProductName:" + Description);
                                                continue;
                                            }
                                            Description += ProductNameValue + " ";
                                            System.out.println("**" + excelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                            continue;
                                        }
                                    }
                                }


                            }
                        }
                    }

                }
                if (!Description.equals("")) {
                    result = Description;
                    System.out.println(result);
                    System.out.println(DES);
                    break;
                }
            }
        }catch (Exception e) {
            e.printStackTrace();
        }

        return result;
    }

}

