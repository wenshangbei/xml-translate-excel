package com.accenture.xmltoexcel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

public class XmlToExcel {

	public static void main(String[] args) {
		

		String xmlPath = "d:/EDI810(HK)-MTL(1224H Stevedoring) MSC 20170629.xml";
		writeExcel( xmlPath);

	
	}
	private static XSSFWorkbook wb;


	 @SuppressWarnings({ "unchecked" })
	    public static Map<String, List<List<String>>> analysisXml(String xmlpath) {
	        // 创建linkedHashmap，为有序map，map和hashMap均为无序的map
	        Map<String, List<List<String>>> sheetMap = new LinkedHashMap<String, List<List<String>>>();
	        // 创建testCaseLists，存放的是testDateLists
	        List<List<List<String>>> testCaseLists = new ArrayList<List<List<String>>>();
	        // 创建SAXReader对象
	        SAXReader reader = new SAXReader();
	        // 读取文件 转换成Document
	        try {
	            Document document = reader.read(new File(xmlpath));
	            // 获取根节点元素对象
	            Element root = document.getRootElement();
	            // 遍历根节点下面的所有子节点
	            List<Element> rootlist = root.elements();

	            // 循环遍历根节点下面的每一子节点
	            for (Element testcase : rootlist) {
	                // 遍历子节点下面的节点
	                List<Element> testCaseList = testcase.elements();
	                // 创建testDateLists，存放的是testDataValue
	                List<List<String>> testDateLists = new ArrayList<List<String>>();
	                // 将testDateLists加进testCaseLists中
	                testCaseLists.add(testDateLists);
	                // 将testCase的name作为key值，testDateLists作为value存进sheetMap中
	                sheetMap.put(testcase.attributeValue("name"), testDateLists);
	                // 循环遍历testCase下面的子节点，testDate
	                for (Element testdata : testCaseList) {
	                    // 遍历testDate节点的属性
	                    List<Attribute> attributes = testdata.attributes();
	                    // 创建testDateValue，将testDateValue add进去testDateLists里面
	                    // ，所以testDateLists的size就是sheet表的行数,testDataValue是sheet表的列数
	                    List<String> testDataValue = new ArrayList<String>();
	                    testDateLists.add(testDataValue);
	                    // 循环遍历属性节点
	                    for (Attribute att : attributes) {
	                        // 将属性值里面的value存进testDateValue里面，属性值里面的key为列的名称
	                        testDataValue.add(att.getValue());
	                    }
	                }
	            }
	        } catch (DocumentException e) {
	            System.out.println("xml file type error! ");
	        }
	        return sheetMap;
	    }

	    /**
	     * @author: 
	     * @function: xml write in excel
	     * @param: path
	     * @exception: IOException
	     * @return: void
	     * @throws IOException
	     * @throws DocumentException
	     */
	    @SuppressWarnings({ "rawtypes", "static-access", "unchecked" })
	    public static void writeExcel(String file) {
	        wb = new XSSFWorkbook();// 创建工作薄
	        // 设置字体
	        XSSFFont font = wb.createFont();
	        font.setFontHeightInPoints((short) 24);
	        font.setFontName("宋体");
	        font.setColor(font.COLOR_NORMAL);
	        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);

	        // 设置单元格样式
	        XSSFCellStyle style = wb.createCellStyle();
	        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
	        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
	        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
	        style.setFont(font);

	        Map<String, List<List<String>>> sheetMap = analysisXml(file);

	        Iterator iterator = sheetMap.entrySet().iterator();
	        while (iterator.hasNext()) {

	            Map.Entry entry = (Map.Entry) iterator.next();
	            // sheetMap的 key为sheet表的名字，value为存的每一行每一列的值
	            XSSFSheet sheet = creatSheetHead(entry.getKey().toString());
	            // 写入单元格
	            List<List<String>> dateLists = (List<List<String>>) entry.getValue();
	            // 行从1开始，0位表头，testDateLists.size为sheet表的行数
	            for (int j = 1; j < dateLists.size() + 1; j++) {
	                XSSFRow row = sheet.createRow(j);
	                // 取list里面的数据从0开始取
	                List<String> dataValue = (List<String>) dateLists.get(j - 1);
	                // 列从0开始，list
	                for (int i = dataValue.size() - 1; i >= 0; i--) {
	                    XSSFCell cell = row.createCell(dataValue.size() - i - 1);
	                    cell.setCellValue(dataValue.get(i));
	                    // 自动设置列宽
	                    sheet.autoSizeColumn(i);
	                }
	            }
	        }
	        createExcel("f://file//模板6.xlsx");
	    }

	    /**
	     * @author:
	     * @function: create excel
	     * @param: path
	     * @exception: IOException
	     * @return: void
	     */
	    public static void createExcel(String path) {
	        ByteArrayOutputStream os = new ByteArrayOutputStream();
	        try {
	            wb.write(os);
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	        byte[] xlsx = os.toByteArray();
	        File file = new File(path);
	        OutputStream out = null;
	        try {
	            out = new FileOutputStream(file);
	            try {
	                out.write(xlsx);
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        }
	    }

	    /**
	     * @author:
	     * @function: create sheet's head
	     * @param: groups
	     * @exception: null
	     * @return: XSSFSheet
	     */
	    private static XSSFSheet creatSheetHead(String sheetName) {
	        // 创建工作表
	        XSSFSheet sheet = wb.createSheet(sheetName);// 创建工作表
	        // 创建第一行第一列，行高为500
	        XSSFRow rowhead = sheet.createRow(0);
	        rowhead.setHeight((short) 500);
	        // 创建表头的内容
	        List<String> sheetHead = new ArrayList<String>();
	        sheetHead.add(0, "description");
	        sheetHead.add(1, "key");
	        sheetHead.add(2, "group");
	        // 写入表头内容
	        for (int j = 0; j < sheetHead.size(); j++) {
	            XSSFCell cellhead = rowhead.createCell(j);
	            cellhead.setCellValue(sheetHead.get(j));
	            // 设置自动列宽
	            sheet.autoSizeColumn(j);
	        }
	        return sheet;
	    }
}

