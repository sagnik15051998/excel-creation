package com.excelcreation.exception;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

@RestControllerAdvice
public class ExcelExceptionHandler {
	
	@ExceptionHandler(Exception.class)
	public ResponseEntity<ErrorResponse> handleException(Exception e) {
		ErrorResponse eresp = new ErrorResponse();
		eresp.setErrorMessage(e.getMessage());
		return new ResponseEntity<>(eresp, HttpStatus.INTERNAL_SERVER_ERROR);
	}
}
