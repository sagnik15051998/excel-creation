package com.excelcreation.dto;

public class ResponseDto {
	
	private String status;
	private String filePath;
	
	public ResponseDto(String status, String filePath) {
		this.status = status;
		this.filePath = filePath;
	}
	
	public String getStatus() {
		return status;
	}
	public String getFilePath() {
		return filePath;
	}
}
