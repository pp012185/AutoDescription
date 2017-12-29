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

public class AutoDescription implements ICustomAction{
    @Override
    public ActionResult doAction(IAgileSession session, INode iNode, IDataObject obj) {

        System.out.println("------ Start ------");

        String filepath = "C:\\ExcelFile\\test_2.xlsx";
        try {
            IChange change =(IChange) obj;
            ITable Affected_tb = change.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);

            Iterator it = Affected_tb.iterator();
            while(it.hasNext()) {
                String result = "";
                IRow row = (IRow) it.next();
                IItem item = (IItem) row.getReferent();     // 直接抓那 row 物件
                String itemNumber = item.getName();
                System.out.println();
                System.out.println("Part Number: " + itemNumber);
/*
                // get subclasss name
                IAgileClass classes = session.getAdminInstance().getAgileClass("10CPU");
                IAgileClass classes2 =item.getAgileClass();

                System.out.println("Class Name : " + classes.getName());
                System.out.println("ID : " + classes.getId());
                System.out.println("API Name : "+classes.getAPIName());

                System.out.println("Class Name : " + classes2.getName());
                System.out.println("ID : " + classes2.getId());
                System.out.println("API Name : "+classes2.getAPIName());
*/
                result+= readExcel(filepath, item, change);
                row.setValue(ChangeConstants.ATT_AFFECTED_ITEMS_ITEM_DESCRIPTION,result);


            }
            System.out.println("------ End ------");
        } catch (APIException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }


        return new ActionResult(0,"Success: ");
    }


    public static String readExcel(String path, IItem item, IChange change)throws Exception {
        String result = "";
        String Description = "";
        try {
            // 用API Name去比對
            // get Api Name of Subclass
            /*
            String tmp =  item.getValue(2020).toString();
            ICell listCell = item.getCell(2020);      // 1081 -> Title block/ Part Type
            IAgileList list = (IAgileList)listCell.getValue();
            String subClass = ((IAgileList)list.getChild(tmp)).getAPIName().toString();
            System.out.println("API subClass:" + subClass);
            */
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
                        if (!"null".equals(excelCell)&&("" + row.getCell(j)).length()!=0) { // excel field not null
                            if(excelCell.equals("end")){
                                break;
                            } else if (excelCell.contains("$")) {            // e.g. $abc -> abc
                                String value = excelCell.substring(1, excelCell.length());
                                Description +=  value + " ";
                            } else if(excelCell.contains("!")){       // 後面不要空格
                                String NewExcelCell= excelCell.substring(0,excelCell.length()-1);
                                String ProductNameValue = item.getValue("Page Three." + NewExcelCell) + "";
                                Description += ProductNameValue;
                                System.out.println("**" + NewExcelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                            }else if (item.getCell("Page Three." + excelCell) == null) { // no this field
                                //System.out.println(item.getCell(1541).getName());
                                System.out.println("no field:" + row.getCell(j));
                                Description += "█" + item + ":no field:" + row.getCell(j) +" ";
                            }  else {// excel field==PLM field
                                IAgileClass cls = item.getAgileClass();
                                IAttribute atr = cls.getAttribute("Page Three." + excelCell);
                                System.out.println(atr.getDataType());
                                if (atr.getDataType()==2) { // 組成為text
                                    String ProductNameValue = item.getValue("Page Three." + excelCell) + "";
                                    Description += ProductNameValue +" " ;
                                    System.out.println("**" + excelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                }else if (atr.getDataType()==4) {// 組成為list
                                                                 // list 加其description 不直接加 name
                                    if(item.getValue("Page Three." + excelCell).toString()=="") {   // 若有對應的list，但是沒值
                                        System.out.println("Field Name:" + excelCell + " -> No Value");
                                        Description += "█" + item + ": Field:" + excelCell + " -> No Value ";
                                        System.out.println("**" + excelCell + "::ProductNameValue:" + " null " + ",ProductName:" + Description);
                                        continue;
                                    }
                                    String tmp2 = item.getValue("Page Three." + excelCell).toString();
                                    ICell listCell2 = item.getCell("Page Three." + excelCell);
                                    IAgileList list2 = (IAgileList)listCell2.getValue();
                                    String ProductNameValue = ((IAgileList)list2.getChild(tmp2)).getDescription();
                                    if(ProductNameValue==null) {
                                        ProductNameValue = "";
                                        Description += ProductNameValue;
                                        System.out.println("**" + excelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                        continue;
                                    }
                                    Description += ProductNameValue + " ";
                                    System.out.println("**" + excelCell + "::ProductNameValue:" + ProductNameValue + ",ProductName:" + Description);
                                }
                            }
                        }
                    }
                    //System.out.println();
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

