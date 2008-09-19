package com.sollink.documentation

import org.codehaus.groovy.scriptom.ActiveXObject

class ExcelGenerator {
	def workbook
	def sheet
	def index
	
	public boolean open_new_excel_doc(){
		def xls = new ActiveXObject('Excel.Application')
		def workbooks = xls.Workbooks
		xls.Visible = true
		workbook = workbooks.Add()
		
		sheet = workbook.Sheets.item(1)
		index = 1
	}
	
	public boolean next_line() {
		index ++
	}
	
	public boolean insert(int row, String value) {
		sheet.Cells(index, row,).value = value
	}
	
	public boolean saveAs(String file_name){
		workbook.SaveAs(new File(file_name).canonicalPath)
	}
	
	public boolean close_excel_doc(){
		workbook.Close()
		xls.Quit()
	}
}