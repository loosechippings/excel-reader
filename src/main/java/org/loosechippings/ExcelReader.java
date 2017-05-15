package org.loosechippings;


import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelReader {
   private Workbook workBook;
   private List<String> headers;
   private Iterator<Row> wbIterator;

   public ExcelReader(File inputFile) throws IOException, InvalidFormatException {
      workBook= WorkbookFactory.create(inputFile);
   }

   public List<String> getSheetNames() {
      List<String> sheetNames=new ArrayList();
      for (Sheet sheet:workBook) {
         sheetNames.add(sheet.getSheetName());
      }
      return sheetNames;
   }

   public List<String> getHeaders(int sheetNumber) {
      if (headers==null) {
         headers=new ArrayList();
         Sheet sheet=workBook.getSheetAt(sheetNumber);
         Row headerRow=sheet.getRow(0);
         for (Cell cell:headerRow) {
            switch (cell.getCellTypeEnum()) {
               case STRING:headers.add(cell.getStringCellValue());
                  break;
               default:
            }
         }
      }
      return headers;
   }

   public String getNextRecord() {
      if (wbIterator==null) {
         wbIterator=workBook.getSheetAt(0).iterator();
         // skip header
         wbIterator.next();
      }
      return format(wbIterator.next());
   }

   private String format(Row row) {
      List<String> cells=new ArrayList<>();
      Cell cell;
      String value;
      String fieldName;
      for (int i=0;i<getHeaders(0).size();i++) {
         cell=row.getCell(i);
         fieldName=getHeaders(0).get(i);
         switch(cell.getCellTypeEnum()) {
            case STRING: value=cell.getStringCellValue();
               break;
            case NUMERIC: value=String.valueOf(cell.getNumericCellValue());
               break;
            default: value="";
         }
         cells.add(String.format("\"%s\":\"%s\"",fieldName,value));
      }
      return "{"+String.join(",",cells)+"}";
   }
}
