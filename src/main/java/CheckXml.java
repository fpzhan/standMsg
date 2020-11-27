import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.dom4j.tree.DefaultAttribute;

import java.io.*;
import java.util.*;

public class CheckXml {
    public static void main(String[] args)throws Exception {
        String topath = Path.rootPath+"\\deal\\finishExcel\\";
        List<String> xmls = new ArrayList<String>();
        List<String> excels = new ArrayList<String>();
        FileUtil.findFileList(new File(Path.rootPath+"\\deal\\xml"),xmls);
        FileUtil.findFileList(new File(Path.rootPath+"\\deal\\excel"),excels);
        System.out.println();
        File file = new File(topath);
        if(!file.exists()){

            file.mkdir();
        }
        File file1 = new File(Path.rootPath+"\\deal\\xml");
        if(!file1.exists()){

            file1.mkdir();
        }
        File file2 = new File(Path.rootPath+"\\deal\\excel");
        if(!file2.exists()){

            file2.mkdir();
        }
        Map<String,String> xmlMap = new HashMap<String, String>();
        for(String xml : xmls){
            String [] xmlNames = xml.split("\\\\");
            String name = xmlNames[xmlNames.length-1];
            name = name.replace(".xml","");
            xmlMap.put(name,xml);
        }
        for(String excel :excels){
            checkExcel(excel,xmlMap,topath);
        }
    }

    private static void checkExcel(String excel,Map<String,String> xmlMap , String topath)throws Exception{

        String [] names = excel.split("\\\\");
        String nameAndType = names[names.length-1];
        String name = nameAndType.replace(".xlsx","").replace(".xls","");
        System.out.println("------------------------------");
        System.out.println("开始处理excel："+nameAndType);
        String xml = xmlMap.get(name);
        SAXReader reader = new SAXReader();
        Document d = reader.read(new File(xml));
        Element root = d.getRootElement();
        List<String> paths = new ArrayList<String>();
        checkXml(root,paths,"");



        Map<String,Set<String>> map = new HashMap<String, Set<String>>();
        InputStream fis = null;
        try {
            fis = new FileInputStream(excel);
            Workbook workbook = null;

            if (excel.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (excel.endsWith(".xls") || excel.endsWith(".et")) {
                workbook = new HSSFWorkbook(fis);
            }
            CellStyle cstyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setColor(Font.COLOR_RED);//字体颜色
//            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            cstyle .setFont(font);
            fis.close();
            int k = workbook.getNumberOfSheets();
            Set<String> columns = new HashSet<String>();
            for (int i = 0; i < k; i++) {
                /* 读EXCEL文字内容 */
                // 获取第一个sheet表，也可使用sheet表名获取
                Sheet sheet = workbook.getSheetAt(i);
// 获取行
                int rowLines = sheet.getLastRowNum();
                int pathLine =0;
                int valueLine=0 ;
                int attrLine=0;
                int mapLine=0;
                int versionLine = 0;
                for(int line=0 ; line<=rowLines;line++) {
                    Row row = sheet.getRow(line);
                    if(row!=null){
                        valueLine++;
                        if (line == 0) {
                            for (int cellLine = 0; cellLine <= row.getLastCellNum(); cellLine++) {
                                Cell cell = row.getCell(cellLine);
                                if (cell != null && "元素路径".equals(cell.getStringCellValue())) {
                                    pathLine = cellLine;
                                }
                                if (cell != null && "属性".equals(cell.getStringCellValue())) {
                                    attrLine = cellLine;
                                }
                                if (cell != null && cell.getStringCellValue().contains("字段")) {
                                    mapLine = cellLine;
                                }
                                if (cell != null && "业务约束".equals(cell.getStringCellValue())) {
                                    versionLine = cellLine;
                                }

                            }
                        } else {
                            if (row.getCell(pathLine) != null && !"".equals(row.getCell(pathLine).getStringCellValue())) {
                                String path = row.getCell(pathLine).getStringCellValue().replace("\\s","");
                                row.getCell(pathLine).setCellValue(path);
                                if(paths.contains(path)){
                                   paths.remove(path);
                                    columns.add(row.getCell(mapLine).getStringCellValue().toUpperCase());
                                }else{
                                    row.getCell(pathLine).setCellStyle(cstyle);
                                }
                            }
                        }
                    }

                }

                for(String path : paths){
                    valueLine++;
                    Row createRow = sheet.createRow(valueLine);
                    String [] pathStrs = path.split("\\.");
                    if(pathStrs[pathStrs.length-1].contains("@")){
                        createRow.createCell(attrLine).setCellValue(pathStrs[pathStrs.length-1].replace("@",""));

                    }else{
                        createRow.createCell(attrLine).setCellValue(pathStrs[pathStrs.length-1]);
                    }
                    createRow.createCell(pathLine).setCellValue(path);
                    if(pathStrs.length>=3 && !columns.contains( pathStrs[pathStrs.length-2].toUpperCase())){
                        createRow.createCell(mapLine).setCellValue( pathStrs[pathStrs.length-2].toUpperCase());
                        columns.add( pathStrs[pathStrs.length-2].toUpperCase());
                    }else if(pathStrs.length>=4 && !columns.contains(pathStrs[pathStrs.length-3].toUpperCase()+"_"+pathStrs[pathStrs.length-2].toUpperCase())){
                        createRow.createCell(mapLine).setCellValue(pathStrs[pathStrs.length-3].toUpperCase()+"_"+pathStrs[pathStrs.length-2].toUpperCase());
                        columns.add(pathStrs[pathStrs.length-3].toUpperCase()+"_"+pathStrs[pathStrs.length-2].toUpperCase());
                    }else if(pathStrs.length>=5 && !columns.contains(pathStrs[pathStrs.length-4].toUpperCase()+"_"+pathStrs[pathStrs.length-3].toUpperCase()+"_"+pathStrs[pathStrs.length-2].toUpperCase())){
                        createRow.createCell(mapLine).setCellValue(pathStrs[pathStrs.length-4].toUpperCase()+"_"+pathStrs[pathStrs.length-3].toUpperCase()+"_"+pathStrs[pathStrs.length-2].toUpperCase());
                        columns.add(pathStrs[pathStrs.length-4].toUpperCase()+"_"+pathStrs[pathStrs.length-3].toUpperCase()+"_"+pathStrs[pathStrs.length-2].toUpperCase());
                    }
                    createRow.createCell(versionLine).setCellValue("1.1");
                }
            }

            FileOutputStream excelFileOutPutStream = new FileOutputStream(topath+"\\"+nameAndType);
            workbook.write(excelFileOutPutStream);
            excelFileOutPutStream.flush();
            excelFileOutPutStream.close();
        }catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.out.println("处理完成excel："+nameAndType);
            System.out.println("------------------------------");

            if (null != fis) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    private static void checkXml(Element element,List<String> paths,String prefixPath){
        List attrs = element.attributes();
        boolean haveCode=false;
        boolean haveDisplayName=false;
        for(Object attr : attrs){
            attr=(DefaultAttribute)attr;
            if("extension".equals(((DefaultAttribute) attr).getName())){
                paths.add(prefixPath+element.getName()+".@extension");
            }
            if("value".equals(((DefaultAttribute) attr).getName())){
                paths.add(prefixPath+element.getName()+".@value");
            }

            if("code".equals(((DefaultAttribute) attr).getName())){
                haveCode=true;
            }

            if("displayName".equals(((DefaultAttribute) attr).getName())){
                haveDisplayName=true;
            }
        }

        if(haveCode && haveDisplayName){
            paths.add(prefixPath+element.getName()+".@code");
            paths.add(prefixPath+element.getName()+".@displayName");
        }

        if(element.elements().size()>0){
            for(Object ele:element.elements()){
                checkXml((Element) ele,paths,prefixPath+element.getName()+".");
            }
        }else if(element.getText()!=null && !"".equals(element.getText())){
            paths.add(prefixPath+element.getName()+".fixedLabelText");
        }
    }



}
