package com.sopra.dtexport;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Main {

  private static final String FOLDER = "d:\\Profiles\\msqueglia\\Desktop\\";
  private static final String XML_NAME = "Dynatrace-exceptions.xml";
  private static final String XLSX_NAME = "Dynatrace-exceptions_";
  private static final String XLSX = ".xlsx";
  private static final String DDMMYYY_HHMMSS = "ddMMyyy_hhmmss";

  public static void main(String[] args) throws IOException, SAXException, ParserConfigurationException {

    File fXmlFile = new File(FOLDER + XML_NAME);
    DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
    Document doc = dBuilder.parse(fXmlFile);

    doc.getDocumentElement().normalize();

    NodeList nList = doc.getElementsByTagName("exceptionrecord");

    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("DT Export");

    createFirstRow(workbook, sheet);

    final XSSFCellStyle style = getStyle(workbook);
    for (int i = 0; i < nList.getLength(); i++) {

      Node nNode = nList.item(i);
      Element eElement = (Element) nNode;

      Row row = sheet.createRow(i + 1);
      final Cell cell0 = row.createCell(0);
      cell0.setCellValue(eElement.getAttribute("exception_class"));
      cell0.setCellStyle(style);

      final Cell cell1 = row.createCell(1);
      cell1.setCellValue(eElement.getAttribute("message"));
      cell1.setCellStyle(style);

      final Cell cell2 = row.createCell(2);
      cell2.setCellValue(eElement.getAttribute("count"));
      cell2.setCellStyle(style);

      final Cell cell3 = row.createCell(3);
      cell3.setCellValue(eElement.getAttribute("throwing_class"));
      cell3.setCellStyle(style);

      final Cell cell4 = row.createCell(4);
      cell4.setCellValue(eElement.getAttribute("throwing_method"));
      cell4.setCellStyle(style);
    }

    sheet.autoSizeColumn(0);
    sheet.autoSizeColumn(1);
    sheet.autoSizeColumn(2);
    sheet.autoSizeColumn(3);
    sheet.autoSizeColumn(4);

    final Date date = new Date();
    final SimpleDateFormat simpleDateFormat = new SimpleDateFormat(DDMMYYY_HHMMSS);
    FileOutputStream outputStream = new FileOutputStream(FOLDER + XLSX_NAME + simpleDateFormat.format(date) + XLSX);
    workbook.write(outputStream);
    System.out.println("END!");
  }

  private static void createFirstRow(XSSFWorkbook workbook, XSSFSheet sheet) {
    final XSSFCellStyle style = workbook.createCellStyle();

    XSSFFont font = workbook.createFont();
    font.setBold(true);
    style.setFont(font);
    style.setAlignment(HorizontalAlignment.CENTER);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);

    Row row = sheet.createRow(0);

    final Cell cell0 = row.createCell(0);
    cell0.setCellValue("exception_class");
    cell0.setCellStyle(style);

    final Cell cell1 = row.createCell(1);
    cell1.setCellValue("message");
    cell1.setCellStyle(style);

    final Cell cell2 = row.createCell(2);
    cell2.setCellValue("count");
    cell2.setCellStyle(style);

    final Cell cell3 = row.createCell(3);
    cell3.setCellValue("throwing_class");
    cell3.setCellStyle(style);

    final Cell cell4 = row.createCell(4);
    cell4.setCellValue("throwing_method");
    cell4.setCellStyle(style);
  }

  private static XSSFCellStyle getStyle(XSSFWorkbook workbook) {
    final XSSFCellStyle style = workbook.createCellStyle();

    XSSFFont font = workbook.createFont();
    style.setFont(font);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    return style;
  }

}