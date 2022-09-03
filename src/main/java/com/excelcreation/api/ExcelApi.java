package com.excelcreation.api;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.excelcreation.dto.ResponseDto;
import com.excelcreation.service.ExcelService;

@Controller
@RequestMapping("/excel")
public class ExcelApi {

	@Autowired
	private ExcelService excelService;
	
	@PostMapping("/create")
	public ResponseEntity<ResponseDto> createExcel() throws Exception {
		return new ResponseEntity<>(excelService.createExcel(), HttpStatus.CREATED);
	}
}
