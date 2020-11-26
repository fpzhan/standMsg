import org.apache.lucene.analysis.Analyzer;
import org.apache.lucene.analysis.TokenStream;
import org.apache.lucene.analysis.tokenattributes.CharTermAttribute;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.SchemaGlobalElement;
import org.wltea.analyzer.lucene.IKAnalyzer;

import java.io.*;
import java.util.*;

public class test {

    private static String [] stopWord = new String[]{"的","其中","root"};

    private static List<String> stopWordList = Arrays.asList(stopWord);


    public static double YUZHI = 0.2 ;

    /**
     * 语义相似度计算 返回百分比
     * @author: Administrator
     * @Date: 2015年1月22日
     * @param T1
     * @param T2
     * @return
     */
    public static double getSimilarity(List<String> T1, List<String> T2) throws Exception {
        int size = 0 , size2 = 0 ;
        if ( T1 != null && ( size = T1.size() ) > 0 && T2 != null && ( size2 = T2.size() ) > 0 ) {

            Map<String, double[]> T = new HashMap<String, double[]>();

            //T1和T2的并集T
            String index = null ;
            for ( int i = 0 ; i < size ; i++ ) {
                index = T1.get(i) ;
                if( index != null){
                    double[] c = T.get(index);
                    c = new double[2];
                    c[0] = 1;	//T1的语义分数Ci
                    c[1] = YUZHI;//T2的语义分数Ci
                    T.put( index, c );
                }
            }

            for ( int i = 0; i < size2 ; i++ ) {
                index = T2.get(i) ;
                if( index != null ){
                    double[] c = T.get( index );
                    if( c != null && c.length == 2 ){
                        c[1] = 1; //T2中也存在，T2的语义分数=1
                    }else {
                        c = new double[2];
                        c[0] = YUZHI; //T1的语义分数Ci
                        c[1] = 1; //T2的语义分数Ci
                        T.put( index , c );
                    }
                }
            }

            //开始计算，百分比
            Iterator<String> it = T.keySet().iterator();
            double s1 = 0 , s2 = 0, Ssum = 0;  //S1、S2
            while( it.hasNext() ){
                double[] c = T.get( it.next() );
                Ssum += c[0]*c[1];
                s1 += c[0]*c[0];
                s2 += c[1]*c[1];
            }
            //百分比
            return Ssum / Math.sqrt( s1*s2 );

        } else {
            throw new Exception("传入参数有问题！");
        }
    }

    private static Map<String,Set<String>> getDocument(String document){

        Map<String,Set<String>> map = new HashMap<String, Set<String>>();
        InputStream fis = null;
        try {
            fis = new FileInputStream(document);
            Workbook workbook = null;
            if (document.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (document.endsWith(".xls") || document.endsWith(".et")) {
                workbook = new HSSFWorkbook(fis);
            }
            fis.close();
            int k = workbook.getNumberOfSheets();
            for(int i = 0 ;i<k;i++){
                /* 读EXCEL文字内容 */
                // 获取第一个sheet表，也可使用sheet表名获取
                Sheet sheet = workbook.getSheetAt(i);
                // 获取行
                int rowLines = sheet.getLastRowNum();
                int columnLine =0;
                int descLine=0;
                for(int line=0 ; line<=rowLines;line++){
                    Row row = sheet.getRow(line);
                    if(line==0){
                        for(int cellLine=0;cellLine<=row.getLastCellNum();cellLine++){
                             Cell cell = row.getCell(cellLine);
                             if(cell!=null && "字段名".equals(cell.getStringCellValue())){
                                 columnLine=cellLine;
                             }
                            if(cell!=null && "说明与描述".equals(cell.getStringCellValue())){
                                descLine=cellLine;
                            }
                        }
                    }else{
                        if(row.getCell(descLine)!=null){
                            String columnName = row.getCell(columnLine).getStringCellValue();
                            if(map.get(columnName)==null){
                                map.put(columnName,new HashSet<String>());
                            }
                            //创建分词对象
                            Analyzer anal=new IKAnalyzer(true);
                            StringReader reader=new StringReader(row.getCell(descLine).getStringCellValue());
                            //分词
                            TokenStream ts=anal.tokenStream("", reader);
                            CharTermAttribute term=ts.getAttribute(CharTermAttribute.class);
                            //遍历分词数据
                            ts.reset();
                            Set<String> tmp = new HashSet<String>();
                            while(ts.incrementToken()){
                                if(stopWordList.contains(term.toString())){
                                    continue;
                                }else{
                                    tmp.add(term.toString());
                                }

                            }
                            if(map.get(columnName).size()==0 || map.get(columnName).size()>tmp.size()){
                                map.put(columnName,tmp);
                            }

                            reader.close();
                        }

                    }
                }
            }


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != fis) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return map;
    }

    private static String getColumnField(String desc,Map<String,Set<String>> docMap)throws Exception{
        //创建分词对象
        Analyzer anal=new IKAnalyzer(true);
        StringReader reader=new StringReader(desc);
        //分词
        TokenStream ts=anal.tokenStream("", reader);
        CharTermAttribute term=ts.getAttribute(CharTermAttribute.class);
        //遍历分词数据
        ts.reset();
        List<String> list = new ArrayList<String>();
        while(ts.incrementToken()){
            if(stopWordList.contains(term.toString())){
                continue;
            }else{
                list.add(term.toString());
            }

        }
        reader.close();

        return getFieldByList(list,docMap);
    }


    private static String getFieldByList(List<String> list,Map<String,Set<String>> docMap)throws Exception{
        if(list.size()==0){
            return "";
        }else{
            Map<Double,String > columns = new HashMap<Double, String>();
            for (Map.Entry<String,Set<String>> entry : docMap.entrySet()){
                double iz = getSimilarity(new ArrayList<String>(entry.getValue()),list);
                columns.put(iz,entry.getKey());
                System.out.println(iz +"---"+entry.getKey()+"---"+entry.getValue());
            }
            Double max = Collections.max(columns.keySet());
            System.out.println("相似值："+max+"相似者："+docMap.get(columns.get(max))+"---"+list);

            return columns.get(max);
        }

    }

    private static void dealData (String data ,Map<String,Set<String>> docMap)throws Exception{
        InputStream fis = null;
        try {
            fis = new FileInputStream(data);
            Workbook workbook = null;
            if (data.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (data.endsWith(".xls") || data.endsWith(".et")) {
                workbook = new HSSFWorkbook(fis);
            }
            fis.close();
            int k = workbook.getNumberOfSheets();
            for(int i = 0 ;i<k;i++){
                /* 读EXCEL文字内容 */
                // 获取第一个sheet表，也可使用sheet表名获取
                Sheet sheet = workbook.getSheetAt(i);
                // 获取行
                int rowLines = sheet.getLastRowNum();
                int pathLine =0;
                int attrLine=0;
                int descLine = 0;
                String rootPath="";
                for(int line=0 ; line<=rowLines;line++){
                    Row row = sheet.getRow(line);
                    if(row!=null){
                        if(line==0){
                            for(int cellLine=0;cellLine<=row.getLastCellNum();cellLine++){
                                Cell cell = row.getCell(cellLine);
                                if(cell!=null && "元素路径".equals(cell.getStringCellValue())){
                                    pathLine=cellLine;
                                }
                                if(cell!=null && "属性".equals(cell.getStringCellValue())){
                                    attrLine=cellLine;
                                }
                                if(cell!=null && "说明与描述".equals(cell.getStringCellValue())){
                                    descLine=cellLine;
                                }
                            }
                            row.createCell(5).setCellValue("字段");
                        }else if(line==1){
                            Cell cell = row.getCell(pathLine);
                            rootPath= cell.getStringCellValue();
                        }else{
                            if(row.getCell(pathLine)!=null){
                                String path=rootPath+row.getCell(pathLine).getStringCellValue();
                                if(row.getCell(attrLine)!=null && !"".equals(row.getCell(attrLine).getStringCellValue())){
                                    path+=(".@"+row.getCell(attrLine).getStringCellValue());
                                }
                                row.getCell(pathLine).setCellValue(path.replace("/","."));
                            }
                            if(row.getCell(descLine)!=null){
                                String desc = row.getCell(descLine).getStringCellValue();
                                String field = getColumnField(desc,docMap);
                                row.createCell(5).setCellValue(field);
                            }
                        }
                    }

                }
                if(sheet!=null && sheet.getRow(1)!=null){

                    sheet.removeRow(sheet.getRow(1));
                }
            }

            FileOutputStream excelFileOutPutStream = new FileOutputStream(data);
            workbook.write(excelFileOutPutStream);
            excelFileOutPutStream.flush();
            excelFileOutPutStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (null != fis) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void main(String[] args)throws Exception {

        List<String> lists = FileUtil.getList("E:\\ideaWorkspace\\testExcle\\files");
        System.out.println();
        for (String value :  lists) {
            List<String> datas = new ArrayList<String>();
            FileUtil.findFileList(new File(value+"\\data"),datas);
            List<String> documents = new ArrayList<String>();
            FileUtil.findFileList(new File(value+"\\document"),documents);
            Map<String,Set<String>> docMap = getDocument(documents.get(0));
            System.out.println(docMap);
            for(String data:datas){
                dealData(data,docMap);

            }
        }
    }
}
