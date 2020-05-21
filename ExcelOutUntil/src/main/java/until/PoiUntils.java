package until;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;


/**
 * @author by
 * @version 1.0
 * @date 2020/1/16 14:01
 */
public class PoiUntils {

    public static Boolean creatExcel(String databaseName, String filepath, PoiUntil poi){
        boolean isSuccess =false;
        List<String> tableNames=poi.getTableNames();
        int tableNum = tableNames.size();
        HSSFWorkbook wb = new HSSFWorkbook();
        //表头样式
        HSSFCellStyle headStyle =poi.getHeaderStyle(wb);
        //非表头样式
        HSSFCellStyle style =poi.getStyle(wb);

        //插入汇总页
        Sheet allSheet =wb.createSheet("目录");
        allSheet.setDefaultColumnWidth(45);
        allSheet.setDefaultRowHeight((short) (100*6));
        Row allSheetRow =allSheet.createRow(0);
        allSheetRow.createCell(1).setCellValue("目录");
        HSSFCellStyle dirStyle =poi.getHyperStyle(wb);
        dirStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dirStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        dirStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        allSheetRow.getCell(1).setCellStyle(dirStyle);
        int index_all=1;
        for (String name:tableNames) {
            allSheetRow=allSheet.createRow(index_all);
            allSheetRow.createCell(1).setCellValue(name);
            allSheetRow.getCell(1).setCellFormula("HYPERLINK(\"#"+"'"+name+"'"+"!A1\",\""+name+"\")");
            allSheetRow.getCell(1).setCellStyle(poi.getHyperStyle(wb));
            index_all++;
        }

        for (int num =0;num<tableNum;num++) {
            String tableName =tableNames.get(num);
            Sheet sheet = wb.createSheet();
            sheet.setDefaultColumnWidth(20);
            sheet.setDefaultRowHeight((short) (100*5));
            Row row = sheet.createRow(0);
            row.createCell(0).setCellStyle(headStyle);
            row.getCell(0).setCellValue("表名");
            row.createCell(1).setCellStyle(headStyle);
            row.getCell(1).setCellValue(tableName);

            row=sheet.createRow(1);
            row.createCell(0).setCellStyle(headStyle);
            row.getCell(0).setCellValue("列名");
            row.createCell(1).setCellStyle(headStyle);
            row.getCell(1).setCellValue("数据类型");
            row.createCell(2).setCellStyle(headStyle);
            row.getCell(2).setCellValue("是否为空");
            row.createCell(3).setCellStyle(headStyle);
            row.getCell(3).setCellValue("默认值");
            row.createCell(4).setCellStyle(headStyle);
            row.getCell(4).setCellValue("主键");
            row.createCell(5).setCellStyle(headStyle);
            row.getCell(5).setCellValue("备注");

            Iterator<Map<String,Object>> it =poi.getColumnStruct(tableName).iterator();
            int index =1;
            while(it.hasNext()){
                index++;
                row=sheet.createRow(index);
                Map<String,Object> data =it.next();
                int i=0;
                for (String key :data.keySet()){
                    HSSFCell cell = (HSSFCell) row.createCell(i);
                    cell.setCellStyle(style);
                    HSSFRichTextString text =new HSSFRichTextString(data.get(key)+"");
                    cell.setCellValue(text);
                    i++;
                }
            }
            Row row1 =sheet.createRow(index+1);
            row1.createCell(0).setCellStyle(headStyle);
            row1.getCell(0).setCellValue("目录");
            row1.getCell(0).setCellFormula("HYPERLINK(\"#目录!A2\",\"目录\")");
            wb.setSheetName(num+1,tableName);
        }


        try{
            FileOutputStream out = new FileOutputStream(new File(filepath+databaseName+".xls"));
            wb.write(out);
            out.close();
            isSuccess = true;
        }catch (IOException e){
            e.printStackTrace();
        }


        return  isSuccess;
    }

    public static void main(String[] args) throws IOException {
        String filepath ="d:/";
       boolean isSuccess =false;
       PoiUntil poi = new PoiUntil();
       isSuccess =creatExcel(poi.getDatabaseName(),filepath,poi);
        System.out.println(isSuccess);
    }
}
