/*
 *abstract: 
 *
 *@author NW
 *
 *Created on 2017-4-28
 *
 */
package com.neville.demo;

import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * XLSXCovertCSV.java 
 * abstract:
 * 
 * hostory:
 * 	 NW 2017-4-28 初始化
 */
public class XLSXCovertCSV {
	private OPCPackage xlsxPackage;
	private int minColumns;
	private PrintStream output;
	private String sheetName;
	
	/**
	 * Parses and shows the content of one sheet using the specified styles and
	 * shared-strings tables.
	 * 
	 * @param styles
	 * @param strings
	 * @param sheetInputStream
	 */
	public List<String[]> processSheet(StylesTable styles, ReadOnlySharedStringsTable strings,
			InputStream sheetInputStream) throws IOException, ParserConfigurationException, SAXException {

		InputSource sheetSource = new InputSource(sheetInputStream);
		SAXParserFactory saxFactory = SAXParserFactory.newInstance();
		SAXParser saxParser = saxFactory.newSAXParser();
		XMLReader sheetParser = saxParser.getXMLReader();
		MyXSSFSheetHandler handler = new MyXSSFSheetHandler(styles, strings, this.minColumns, this.output,this.minColumns);
		sheetParser.setContentHandler(handler);
		sheetParser.parse(sheetSource);
		return handler.getRows();
	}
	
	/**
	 * 初始化这个处理程序 将
	 * 
	 * @throws IOException
	 * @throws OpenXML4JException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 */
	public List<String[]> process() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {

		ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
		XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
		List<String[]> list = null;
		StylesTable styles = xssfReader.getStylesTable();
		XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		int index = 0;
		while(iter.hasNext()) {
			InputStream stream = iter.next();
			String sheetNameTemp = iter.getSheetName();
			if (this.sheetName.equals(sheetNameTemp)) {
				list = processSheet(styles, strings, stream);
				stream.close();
				++index;
			}
		}
		return list;
	}
	
	/**
	 * @param xlsxPackage
	 * @param minColumns
	 * @param output
	 * @param sheetName
	 */
	public XLSXCovertCSV(OPCPackage xlsxPackage, int minColumns, PrintStream output, String sheetName) {
		super();
		this.xlsxPackage = xlsxPackage;
		this.minColumns = minColumns;
		this.output = output;
		this.sheetName = sheetName;
	}
	/**
	 * @return the xlsxPackage
	 */
	public OPCPackage getXlsxPackage() {
		return xlsxPackage;
	}
	/**
	 * @param xlsxPackage the xlsxPackage to set
	 */
	public void setXlsxPackage(OPCPackage xlsxPackage) {
		this.xlsxPackage = xlsxPackage;
	}
	/**
	 * @return the minColumns
	 */
	public int getMinColumns() {
		return minColumns;
	}
	/**
	 * @param minColumns the minColumns to set
	 */
	public void setMinColumns(int minColumns) {
		this.minColumns = minColumns;
	}
	/**
	 * @return the output
	 */
	public PrintStream getOutput() {
		return output;
	}
	/**
	 * @param output the output to set
	 */
	public void setOutput(PrintStream output) {
		this.output = output;
	}
	/**
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}
	/**
	 * @param sheetName the sheetName to set
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
}
