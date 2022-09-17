import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MetaData2 {
    public static void main(String args[]) throws IOException {
        Map<String, List<String>> hm = new HashMap<String, List<String>>();
        List<String> values = new ArrayList<String>();
        String excelFilePath1 = "./FMTC_Results.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath1);
        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
        XSSFSheet Links = workbook1.getSheetAt(0);
        for (int i = 1; i <= Links.getLastRowNum(); i++) {
        try {
            XSSFRow Row = Links.getRow(i);
            XSSFCell Merchant = Row.getCell(1);
            String Merchant_Link = Merchant.getStringCellValue();
            Document doc = Jsoup.connect(Merchant_Link).userAgent("Chrome/104.0.5112.102").timeout(5000).get();
            System.out.printf("Title: %s\n", doc.title());
            String title = doc.title();
            String des = doc.select("meta[name=description]").attr("content");
            System.out.println(des);
            values.add(title);
            values.add(des);
            hm.put(Merchant_Link, values);
            values = new ArrayList<>();
        } catch (IOException e) {
            e.printStackTrace();
            continue;
        }
        }
        int rowNo=0;
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("sheet1");
        XSSFRow row= null;
        for(HashMap.Entry<String,List<String>> entry:hm.entrySet()) {
            row=sheet.createRow(rowNo++);
            row.createCell(0).setCellValue((String)entry.getKey());
            for (int i = 0; i < entry.getValue().size(); i++) {
                row.createCell(1+i).setCellValue(entry.getValue().get(i));
            }
        }
        File filePath = new File("./FMTC_Results1.xlsx");
        if(!filePath.exists()){
            filePath.createNewFile();
        }

        FileOutputStream file = new FileOutputStream(new File("./FMTC_Results1.xlsx"));
        workbook.write(file);
        file.close();
    }
}
