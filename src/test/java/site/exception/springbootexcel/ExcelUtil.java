package site.exception.springbootexcel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Font;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.metadata.TableStyle;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.assertj.core.util.Lists;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

    @Test
    public void testexcel() throws Exception {
//        List list = new ArrayList();
//        list.add("haha");
//        list.add("xixi");
//        list.add("heihei");
//        list.add("luelue");
//        List dataList = new ArrayList();
//        List listData = new ArrayList();
//        listData.add("1");
//        listData.add("2");
//        listData.add("3");
//        listData.add("4");
//        dataList.add(listData);
//        dataList.add(listData);
//        writeChongqingExcel(list,dataList,"C:/Users/Sxh/Desktop/test/chongqing.xlsx",6,"");

//        String serial_number = "11111111111";
//        String temp = serial_number.substring(4, serial_number.length() - 4);
//        StringBuffer buffer = new StringBuffer();
//        buffer.append(serial_number.substring(0, 3));
//        buffer.append("****");
//        buffer.append(serial_number.substring(serial_number.length() - 4, serial_number.length()));
//        serial_number = buffer.toString();
//        System.out.println(serial_number);
        double remainBalance = 10000;
        if (remainBalance>-100){
            System.out.println("aaaa");
            System.out.println(remainBalance);
        }
    }

    @Test
    public void transTelephone(){
        String balance = null;
        BigDecimal number = new BigDecimal(-19999.90);
        if (number.intValue()<1000&&number.intValue()>-100){
            balance = number.setScale(1,BigDecimal.ROUND_HALF_UP).toString();
        }else if (number.intValue()<10000&&number.intValue()>1000){
            balance = number.setScale(1,BigDecimal.ROUND_HALF_UP).toString();
        }else if(number.intValue()<-100&&number.intValue()>-1000){
            balance = number.setScale(1,BigDecimal.ROUND_HALF_UP).toString();
        }else if(number.intValue()<-1000&&number.intValue()>-10000){
            balance = number.setScale(0,BigDecimal.ROUND_HALF_UP).toString();
        }else{
//            balance = new BigDecimal(number.intValue()*0.0001).setScale(1,BigDecimal.ROUND_HALF_UP).toString() + "万";
            balance = new BigDecimal(number.intValue()*0.0001).setScale(1,BigDecimal.ROUND_HALF_UP).toString();
        }
        System.out.println(number);
        System.out.println(balance);
        System.out.println(new BigDecimal(balance));
    }

    /**
     *
     * @param list 数据名称集合
     * @param datalist 数据集合
     * @param route 路径
     * @throws Exception
     */
//    @Test
    public void writeChongqingExcel(List list,List datalist,String route,int colunm,String sheetName) throws Exception {
//        OutputStream out = new FileOutputStream("C:/Users/Sxh/Desktop/test/chongqing.xlsx");
        //下载路径
        OutputStream out = new FileOutputStream(route);
        ExcelWriter writer = EasyExcelFactory.getWriter(out);

        Sheet sheet1 = new Sheet(1, 0);
        //sheet标题
        if (StringUtils.isNotBlank(sheetName)){
            sheet1.setSheetName(sheetName);
        }
        // 设置列的宽度
//        Map columnWidth = new HashMap();
//        columnWidth.put(0, 10);
//        sheet1.setColumnWidthMap(columnWidth);
//        sheet1.setAutoWidth(true);
        Table table1 = new Table(1);

        // 自定义表格样式
        table1.setTableStyle(createChongqingTableStyle());
//        // 无注解的模式，动态添加表头
        table1.setHead(createTestListStringHead(list,colunm));

        //根据数据集合获取每条list中的数据
        List<List> listReturn = new ArrayList();
        for (int i = 0;i<datalist.size();i++){
            listReturn.add((List) datalist.get(i));
        }
        // 内容
        writer.write1(createChongqingDataList(listReturn), sheet1, table1);

        // 合并单元格 todo
        writer.merge(12, 12, 0, 11);


        writer.finish();
        out.close();
    }

    /**
     * 无注解的实体类
     *
     * @return
     */
    private List<List<Object>> createChongqingDataList(List<List> list) {
        List<List<Object>> datas = Lists.newArrayList();
        for (int i = 0; i< list.size();i++){
            List dataList = list.get(i);
            List returnList = new ArrayList();
            for (int j = 0; j<dataList.size();j++){
                returnList.add(dataList.get(j));
            }
            datas.add(returnList);
        }
        return datas;
    }

    /**
     *
     * @param list 表头集合
     * @param column 表头名称需要占用的列数
     * @return
     */
    public static List<List<String>> createTestListStringHead(List list,int column){
        // 模型上没有注解，表头数据动态传入
        List<List<String>> head = new ArrayList<List<String>>();
//        if (list.size()==0||list==null){
//            return null;
//        }
        //获取表头名称
        for (int i=0;i<list.size();i++){
            List list1 = new ArrayList();
            //获取表头名称需要占用的列数
            for(int j=0;j<column;j++){
                list1.add(list.get(i));
            }
            head.add(list1);
        }
        return head;
    }

    public static TableStyle createChongqingTableStyle() {
        // 表头样式
        TableStyle tableStyle = new TableStyle();
        Font headFont = new Font();
        headFont.setBold(true);
        headFont.setFontHeightInPoints((short)12);
        headFont.setFontName("微软雅黑");
        tableStyle.setTableHeadFont(headFont);
        // 内容背景色
        tableStyle.setTableHeadBackGroundColor(IndexedColors.ROSE);

        // 内容样式
        Font contentFont = new Font();
        contentFont.setFontHeightInPoints((short)12);
        contentFont.setFontName("微软雅黑");
        tableStyle.setTableContentFont(contentFont);
        // 内容背景色
        tableStyle.setTableContentBackGroundColor(IndexedColors.WHITE);

        return tableStyle;
    }
}
