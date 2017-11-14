package com.lianglianglee;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CsvToExcel {
  public static void main(String[] args) throws FileNotFoundException, IOException {
    System.out.println(new File("cityair.csv").getAbsolutePath());
    csvToExcel(new File("cityair.csv"), new File("test.xls"));
  }

  public static File csvToExcel(File csvfile, File excelFile)
      throws FileNotFoundException, IOException {
    DataInputStream in = new DataInputStream(new FileInputStream(csvfile));
    // 编码
    byte[] b = new byte[3];
    in.read(b);
    in.close();
    String code;
    if (b[0] == -17 && b[1] == -69 && b[2] == -65) {
      code = "UTF-8";
    } else {
      code = "GBK";
    }
    in.close();
    in = new DataInputStream(new FileInputStream(csvfile));
    BufferedReader br = new BufferedReader(new InputStreamReader(in, code));
    CSVParser records = CSVFormat.EXCEL.parse(br);
    List<CSVRecord> iterator = records.getRecords();
    // 先处理表头
    CSVRecord titleRecord = iterator.get(0);
    int titleCount = titleRecord.size();
    List<String> list = new ArrayList<String>(titleCount);
    for (int j = 0; j < titleCount; j++) {
      list.add(titleRecord.get(j));
    }
    // 开始处理value
    List<String>[] values = new ArrayList[titleCount];
    // 初始化数组
    for (int i = 0; i < values.length; i++) {
      List<String> value = new ArrayList<String>();
      values[i] = value;
    }
    // 循环读取value
    for (int i = 1; i < iterator.size(); i++) {
      CSVRecord csvRecord = iterator.get(i);
      int count = csvRecord.size();
      for (int j = 0; j < count; j++) {
        List<String> value = values[j];
        value.add(csvRecord.get(j));
        values[j] = value;
      }
    }
    // 删除源文件
    csvfile.delete();
    // 组建一个表格需要的实体
    List<DataEntity> dataEntities = new ArrayList<DataEntity>(titleCount);
    for (int i = 0; i < titleCount; i++) {
      DataEntity entity = new DataEntity();
      entity.setTitle(list.get(i));
      entity.setValue(values[i]);
      dataEntities.add(entity);
    }
    if(!excelFile.exists()){
      excelFile.createNewFile();
    }
    createExcel(excelFile, dataEntities);
    return excelFile;
  }

  public static void createExcel(File file, List<DataEntity> entities) {
    Workbook xwb = new HSSFWorkbook();
    Sheet xsheet = xwb.createSheet(); // 获取excel表的第一个sheet
    Row titleRow = xsheet.createRow(0);
    for (int i = 0; i < entities.size(); i++) {
      DataEntity entity = entities.get(i);
      Cell titCell = titleRow.createCell(i);
      titCell.setCellValue(entity.getTitle());
      List<String> values = entity.getValue();
      for (int j = 0; j < values.size(); j++) {
        Row valueRow = null;
        if (xsheet.getLastRowNum() > j) {
          valueRow = xsheet.getRow(j + 1);
        } else {
          valueRow = xsheet.createRow(j + 1);
        }

        Cell valueCell = valueRow.createCell(i);
        valueCell.setCellValue(values.get(j));
      }
    }
    try {
      if (file.exists()) {
        file.createNewFile();
      }
      FileOutputStream fout = new FileOutputStream(file);
      xwb.write(fout);
      fout.close();
    } catch (Exception e) {
      e.printStackTrace();
    }

  }

}
