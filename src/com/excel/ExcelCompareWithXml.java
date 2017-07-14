package com.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

public class ExcelCompareWithXml {

	public List<Map<String,String>> readExcel(String path) throws IOException {
		InputStream is = new FileInputStream(path);
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
		List<Map<String, String>> list = new ArrayList<>();
		for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
			XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
			if (xssfSheet == null||numSheet==5) {
				continue;
			}
			Map<String, String> map = new HashMap<>();
			for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
				XSSFRow xssfRow = xssfSheet.getRow(rowNum);
				if (xssfRow != null) {
					String value = xssfRow.getCell(0).getStringCellValue();
					String key = xssfRow.getCell(1).getStringCellValue();
					if (!"string".equals(value)) {
						map.put(key, value);
					}
				}
			}
			list.add(map);
		}
		return list;
	}

	public List<Map<String,String>> readXml(String path, String lang) throws Exception {
		String[] paths={path + "//ApkManager//" + lang,path + "//CtvMenu//" + lang,path + "//FileManager//" + lang,path + "//Launcher//" + lang,path + "//NetworkSettings//" + lang,path + "//Settings//" + lang};
		List<Map<String,String>> list=new ArrayList<>();
		for(String p:paths){
			Map<String, String> map = new HashMap<>();
			SAXReader reader = new SAXReader();
			InputStream in = new FileInputStream(p);
			Document document = reader.read(in);
			Element root = document.getRootElement();
			for(Iterator iter = root.elementIterator(); iter.hasNext();){
				Element element = (Element) iter.next();
			    Attribute attr=element.attribute("name");
			   map.put(attr.getText(),element.getText());
			}
			list.add(map);
			in.close();
		}
		return list;
	}

	public static void main(String[] args) throws Exception {
		String excelPath = "E://parse//excel//Spanish_ES_1_pt.xlsx";
		String xmlPath = "E://parse//xml";
		String lang = "strings_pt.xml";
		ExcelCompareWithXml o = new ExcelCompareWithXml();
		List<Map<String,String>> excelList=o.readExcel(excelPath);
		List<Map<String,String>> xmlList=o.readXml(xmlPath, lang);
		String[] types={"ApkManager","CtvMenu","FileManager","Launcher","NetworkSettings","Settings"};
		for(int i=0;i<excelList.size();i++){
			Map<String,String> excelMap=excelList.get(i);
			Map<String,String> xmlMap=xmlList.get(i);
			
			Set<String> keys=excelMap.keySet();
			for(String key:keys){
				if(!excelMap.get(key).equals(xmlMap.get(key))){
					System.out.println(types[i]+":"+key+"的值excel与xml不同"+",excel是"+excelMap.get(key)+",xml是"+xmlMap.get(key));
				}
			}
			System.out.println("=================================================================================");
		}

	}
}
