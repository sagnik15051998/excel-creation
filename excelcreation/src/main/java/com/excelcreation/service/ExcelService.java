package com.excelcreation.service;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import com.excelcreation.dto.ResponseDto;

@Service
public class ExcelService {
	
	@Value("${spring.datasource.url}")
	private String db_url;
	
	@Value("${spring.datasource.username}")
	private String db_user;
	
	@Value("${spring.datasource.password}")
	private String db_pass;
	
	public ResponseDto createExcel() throws Exception {
		Connection conn = DriverManager.getConnection(db_url, db_user, db_pass);
		Statement st = conn.createStatement();
		String queryString = "select * from employees";
		ResultSet rs = st.executeQuery(queryString);
		
		ResultSetMetaData metaResult = rs.getMetaData();
		int colCount = metaResult.getColumnCount();
		int rowCount = 0;
		List<String> columnNames = new ArrayList<>();
		for(int i=1;i<=colCount;i++) {
			columnNames.add(metaResult.getColumnName(i));
		}
		
		String filepath = getFilePath();
		FileOutputStream fos = new FileOutputStream(filepath);
		Workbook wb = new SXSSFWorkbook();
		Sheet sheet = (SXSSFSheet) wb.createSheet("Data");
		CellStyle cs = wb.createCellStyle();
		Font f = wb.createFont();
		f.setBold(true);
		cs.setFont(f);
		cs.setAlignment((short) HorizontalAlignment.CENTER.ordinal());
		cs.setVerticalAlignment((short) VerticalAlignment.CENTER.ordinal());
		Row headRow = sheet.createRow(++rowCount);
		fillHead(headRow, columnNames, cs);
		while(rs.next()) {
			Row dataRow = sheet.createRow(++rowCount);
			fillData(dataRow, rs, colCount);
		}
		st.close();
		conn.close();
		wb.write(fos);
		wb.close();
		fos.close();
		
		return new ResponseDto("Successfully Created", filepath);
	}
	
	private String getFilePath() {
		LocalDateTime currentTime = LocalDateTime.now();
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
		String timing = currentTime.format(dtf);
		String filename = "src/target-for-excel/DATA-FILE-"+timing+".xlsx";
		return filename;
	}
	
	private void fillHead(Row row, List<String> columns, CellStyle cs) {
		int l = columns.size();
		for(int i=1;i<=l;i++) {
			row.createCell(i).setCellType(Cell.CELL_TYPE_STRING);
			row.getCell(i).setCellStyle(cs);
			row.getCell(i).setCellValue(columns.get(i-1));
		}
	}
	
	private void fillData(Row row, ResultSet rs, int colcount) throws Exception {
		int i=1;
		while(i<=colcount) {
			Object val = rs.getObject(i);
			if(val instanceof Long || val instanceof Integer) {
				row.createCell(i).setCellType(Cell.CELL_TYPE_NUMERIC);
				row.getCell(i).setCellValue(Long.valueOf(String.valueOf(val)));
			}
			else {
				row.createCell(i).setCellType(Cell.CELL_TYPE_STRING);
				row.getCell(i).setCellValue(String.valueOf(val));
			}
			i++;
		}
		
	}
	
}
