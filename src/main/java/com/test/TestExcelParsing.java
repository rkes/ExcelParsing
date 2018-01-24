package com.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.xml.sax.SAXException;

public class TestExcelParsing {

	public static void main(String[] args) throws IOException, OpenXML4JException, SAXException {
		FileInputStream inputStream = new FileInputStream(new ClassPathResource("import.xlsx").getFile());
		Set<Map> res = new LinkedHashSet<Map>();
		try (XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream)) {
			XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(2);
			Iterator<Row> rowIterator = xssfSheet.rowIterator();
			XSSFRow header = (XSSFRow) rowIterator.next();// This will Skip
			List<String> headerNames = new ArrayList<String>();
			for (int i = 0; i < header.getLastCellNum(); i++)
				headerNames.add(header.getCell(i).getStringCellValue());
			Integer cellNums = headerNames.size();
			boolean endOfSheet = false;
			while (rowIterator.hasNext() && !endOfSheet) {
				XSSFRow row = (XSSFRow) rowIterator.next();
				HashMap<String, String> resMap = new HashMap();
				for (int i = 0; i < row.getLastCellNum(); i++) {
					if (row.getCell(i).getCellTypeEnum().compareTo(CellType.BLANK) == 0
							|| cellNums != row.getLastCellNum()) {
						endOfSheet = true;
						break;
					}
					if (getCellValue(row.getCell(i)).trim().equals(""))
						continue;
					resMap.put(headerNames.get(i), getCellValue(row.getCell(i)));
				}
				if (!endOfSheet)
					res.add(resMap);
			}
		}
		System.out.println(res);
	}

	private static String getCellValue(XSSFCell cell) {
		switch (cell.getCellTypeEnum()) {
			case NUMERIC:
				return String.valueOf(cell.getNumericCellValue());
			case STRING:
				return cell.getStringCellValue();
			case FORMULA:
				return String.valueOf(cell.getCellFormula());
			case BLANK:
				return "";
			case BOOLEAN:
				return String.valueOf(cell.getBooleanCellValue());
			case ERROR:
				return String.valueOf(cell.getErrorCellString());
			default:
				throw new IllegalArgumentException("Invalid cell type " + cell.getCellTypeEnum());
		}
	}
}
