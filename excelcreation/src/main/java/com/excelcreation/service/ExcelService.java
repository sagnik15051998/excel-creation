package com.excelcreation.service;

import java.io.FileOutputStream;
import java.net.URI;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

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
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import com.excelcreation.dto.ResponseDto;
import com.fasterxml.jackson.databind.ObjectMapper;

@Service
public class ExcelService {
	
	@Value("${employee-uri}")
	private String employeeUrl;
	
	@SuppressWarnings("unchecked")
	public ResponseDto createExcel() throws Exception {

		RestTemplate template = new RestTemplate();
		String filepath = getFilePath();

		String jsons = template.getForObject(new URI(employeeUrl), String.class);
		JSONArray jsona = new JSONArray(jsons);
		JSONObject jsono = jsona.getJSONObject(0);
		
		Map<String, String> map = new ObjectMapper().readValue(jsono.toString(), Map.class);
		List<String> columnNames = map.keySet().stream().collect(Collectors.toList());
		columnNames.sort(Comparator.comparing((String name) -> name.length()));
		
		FileOutputStream fos = new FileOutputStream(filepath);
		Workbook wb = new SXSSFWorkbook();
		Sheet sheet = (SXSSFSheet) wb.createSheet("Data");
		CellStyle cs = wb.createCellStyle();
		Font f = wb.createFont();
		f.setBold(true);
		cs.setFont(f);
		cs.setAlignment((short) HorizontalAlignment.CENTER.ordinal());
		cs.setVerticalAlignment((short) VerticalAlignment.CENTER.ordinal());
		
		int rowCount=0, colCount = jsona.length();
		Row headRow = sheet.createRow(++rowCount);
		
		fillHead(headRow, columnNames, cs);
		
		for(int i=0; i<jsona.length(); i++) {
			Row dataRow = sheet.createRow(++rowCount);
			fillData(dataRow, jsona.getJSONObject(i), colCount, columnNames);
		}

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
	
	private void fillData(Row row, JSONObject jo, int colcount, List<String> columnNames) throws Exception {
		int i=1;
		while(i<=colcount) {
			Object val = jo.get(columnNames.get(i-1));
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
