package com.accenture.xmltoexcel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.activation.UnsupportedDataTypeException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
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
import org.dom4j.QName;
import org.dom4j.io.SAXReader;

public class XmlToExcel2 {

	public static void main(String[] args) {

		String xmlPath = "D:/workspace/EDI810(HK)-MTL(1224H Stevedoring) MSC 20170629.xml";
		// writeExcel( xmlPath);
	/*	Map<String, List<List<String>>> analysisXml = analysisXml(xmlPath);
		System.out.println(analysisXml);*/
		Map<String, List<List<String>>> xml = analysisXml(xmlPath);
		udateExecl(xml);

	}

	private static XSSFWorkbook wb;

	@SuppressWarnings({ "unchecked" })
	public static Map<String, List<List<String>>> analysisXml(String xmlpath) {
		// 创建linkedHashmap，为有序map，map和hashMap均为无序的map
		Map<String, List<List<String>>> worksheet = new LinkedHashMap<String, List<List<String>>>();
		// 创建SAXReader对象
		SAXReader reader = new SAXReader();
		// 读取文件 转换成Document
		try {
			Document document = reader.read(new File(xmlpath));
			// 获取根节点元素对象
			Element root = document.getRootElement();
			// 遍历根节点下面的所有子节点

			List<Element> elements = root.elements("Worksheet");

			for (Element element : elements) {
				System.out.println(element.attributeValue("Name"));
				List<Element> elementsTable = element.elements("Table");

				for (Element element2 : elementsTable) {
					int i = 0;

					List<Element> elements2 = element2.elements("Row");
					ArrayList<List<String>> arrayListRow = new ArrayList<>();
					for (Element element3 : elements2) {
						i++;
						if (i == 1)
							continue;
						ArrayList<String> arrayListCell = new ArrayList<>();
						List<Element> elements3 = element3.elements("Cell");
						for (Element element4 : elements3) {

							arrayListCell.add(element4.getStringValue());
						}
						arrayListRow.add(arrayListCell);
						worksheet.put(element.attributeValue("Name"), arrayListRow);
					}
				}

			}

		} catch (DocumentException e) {
			System.out.println("xml file type error! ");
		}
		return worksheet;
	}

	public static void udateExecl(Map xml) {
		try {
			xml = (LinkedHashMap<String, List<List<String>>>) xml;
			System.out.println(xml);
			File file = new File("D:\\workspace\\MSC_1224H.xls");
			// 传入的文件
			FileInputStream fileInput = new FileInputStream(file);
			// poi包下的类读取excel文件
			POIFSFileSystem ts = new POIFSFileSystem(fileInput);
			// 创建一个webbook，对应一个Excel文件
			HSSFWorkbook workbook = new HSSFWorkbook(ts);
			// 对应Excel文件中的sheet，0代表第一个
			int numberOfSheets = workbook.getNumberOfSheets();
			
			System.out.println(numberOfSheets);
			
			for(int i = 0; i < numberOfSheets; i++) {
				
				HSSFSheet sheet = workbook.getSheetAt(i);
				if (sheet != null) {
					List<List<String>> rows = (List<List<String>>) xml.get(sheet.getSheetName());
					int rowN = 1;
					for (List<String> row : rows) {
						int cellN = 0;
						HSSFRow rowSheet = sheet.getRow(rowN);
						if (rowSheet == null) {
							rowSheet = sheet.createRow(rowN);
						}
						rowN ++;
						for (String cell : row) {
							HSSFCell cellSheet =  rowSheet.getCell(cellN);
							if(cellSheet == null){
								cellSheet = rowSheet.createCell(cellN);
							}
							cellSheet.setCellValue(cell);
							cellN ++;
						}
					}
				}
				
			}
		
			
			
		
			FileOutputStream os = new FileOutputStream("D:\\workspace\\updeData\\MSC_1224H.xls");
			os.flush();
			// 将Excel写出
			workbook.write(os);
			// 关闭流
			fileInput.close();
			os.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

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
