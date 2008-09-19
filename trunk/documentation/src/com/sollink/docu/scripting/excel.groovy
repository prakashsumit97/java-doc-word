import org.codehaus.groovy.scriptom.ActiveXObject

class UseCase {
	public String 세부모듈명;
	public String 세부모듈식별자;
	public String 작성자;
	public String 유즈케이스명;
	public String 유즈케이스식별자
	public String 작성일시;
	public String 관련액터;
	public String 설명;
	public String 사전조건;
	public String 사후조건;
	
	public Vector<String> 기본흐름_단계설명 = new Vector<String>();
	
	public Vector<String> 대안흐름_번호 = new Vector<String>();
	public Vector<String> 대안흐름_조건 = new Vector<String>();
	public Vector<String> 대안흐름_단계설명 = new Vector<String>();
	
	def print() {
		println "Use Case"
		println 세부모듈명
		println 세부모듈식별자
		println 작성자
		println 유즈케이스명
		println 유즈케이스식별자
		println 작성일시
		println 관련액터
		println 설명
		println 사전조건
		println 사후조건
		println 기본흐름_단계설명
		println 대안흐름_번호
		println 대안흐름_조건
		println 대안흐름_단계설명
	}
}

def add_usecase_table(uc, word) {
	def doc = word.ActiveDocument
	
	word.Selection.MoveDown(5, 3)
	word.Selection.TypeText("\n")
	def range = word.Selection.Range;
	
	//
	range.Select()
	int length = 8 + uc.기본흐름_단계설명.size() + 2 +  uc.대안흐름_단계설명.size() 
	def table = doc.Tables.Add(range,length,6)
	
	// 1열
	table.Cell(1,1).Range.Text = "세부 모듈 명"
	table.Cell(1,1).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(1,2).Range.Text = uc.세부모듈명
	
	table.Cell(1,3).Range.Text = "세부 모듈 식별자"
	table.Cell(1,3).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(1,4).Range.Text = uc.세부모듈식별자
	
	table.Cell(1,5).Range.Text = "작성자"
	table.Cell(1,5).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(1,6).Range.Text = uc.작성자
	
	// 2열
	table.Cell(2,1).Range.Text = "유즈케이스 명"
	table.Cell(2,1).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(2,2).Range.Text = uc.유즈케이스명
	
	table.Cell(2,3).Range.Text = "유즈케이스 식별자"
	table.Cell(2,3).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(2,4).Range.Text = uc.유즈케이스식별자
	
	table.Cell(2,5).Range.Text = "작성 일시"
	table.Cell(2,5).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(2,6).Range.Text = uc.작성일시
	
	// 3열 ~ 6열
	table.Cell(3,1).Range.Text = "관련액터"
	table.Cell(3,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(4,1).Range.Text = "설명"
	table.Cell(4,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(5,1).Range.Text = "사전조건"
	table.Cell(5,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(6,1).Range.Text = "사후조건"
	table.Cell(6,1).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(3,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(3,2).Range.Text = uc.관련액터
	
	table.Cell(4,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(4,2).Range.Text = uc.설명
	
	table.Cell(5,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(5,2).Range.Text = uc.사전조건
	
	table.Cell(6,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(6,2).Range.Text = uc.사후조건
	
	// 기본 흐름 (7, 8)
	table.Cell(7,1).Select()
	word.Selection.MoveRight(1, 5, 1)
	word.Selection.Cells.Merge()
	table.Cell(7,1).Range.Text = "기본 흐름"
	table.Cell(7,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(7,1).Range.ParagraphFormat.Alignment = 1
	
	
	table.Cell(8,1).Range.Text = "번호"
	table.Cell(8,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(8,1).Range.ParagraphFormat.Alignment = 1
	
	table.Cell(8,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(8,2).Range.Text = "단계설명"
	table.Cell(8,2).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(8,2).Range.ParagraphFormat.Alignment = 1
	
	//
	int row = 9
	for(int i=0; i<uc.기본흐름_단계설명.size(); i++) {
		table.Cell(row, 1).Range.Text = "${i+1}"
		table.Cell(row,2).Select()
		word.Selection.MoveRight(1, 4, 1)
		word.Selection.Cells.Merge()
		table.Cell(row, 2).Range.Text = uc.기본흐름_단계설명.get(i)
		
		row ++
	}
	
	// 대안 흐름
	
	table.Cell(row,1).Select()
	word.Selection.MoveRight(1, 5, 1)
	word.Selection.Cells.Merge()
	table.Cell(row,1).Range.Text = "대안 흐름"
	table.Cell(row,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,1).Range.ParagraphFormat.Alignment = 1
	
	row ++
	
	table.Cell(row,1).Range.Text = "번호"
	table.Cell(row,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,1).Range.ParagraphFormat.Alignment = 1
	
	table.Cell(row,2).Range.Text = "조건"
	table.Cell(row,2).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,2).Range.ParagraphFormat.Alignment = 1
		
	table.Cell(row,3).Select()
	word.Selection.MoveRight(1, 3, 1)
	word.Selection.Cells.Merge()
	table.Cell(row,3).Range.Text = "단계설명"
	table.Cell(row,3).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,3).Range.ParagraphFormat.Alignment = 1
	
	row ++
	
	//
	for(int i=0; i<uc.대안흐름_단계설명.size(); i++) {
		table.Cell(row, 1).Range.Text = uc.대안흐름_번호.get(i)
		table.Cell(row, 2).Range.Text = uc.대안흐름_조건.get(i)
		
		table.Cell(row, 3).Select()
		word.Selection.MoveRight(1, 3, 1)
		word.Selection.Cells.Merge()
		table.Cell(row, 3).Range.Text = uc.대안흐름_단계설명.get(i)
		
		row ++
	}

}



def String get_value(sheet, range) {
	try {
		String value = ""
		for(c in sheet.Range(range).Value.value) {
			value += c
		}
		return value
	}
	catch(Exception e) {
		return ""
	}
}

def UseCase print_sheet(sheet) {
	def 세부모듈명_value = "B1"
	def 세부모듈식별자_value = "D1"
	def 작성자_value = "F1"
			
	def 유즈케이스명_value = "B2"
	def 유즈케이스식별자_value = "D2"
	def 작성일시_value = "F2"
	
	def 관련액터_value = "B3"
	def 설명_value = "B4"
	def 사전조건_value = "B5"
	def 사후조건_value = "B6"
		
	UseCase uc = new UseCase()
	
	// header
	uc.세부모듈명 = get_value(sheet, 세부모듈명_value)
	uc.세부모듈식별자 = get_value(sheet, 세부모듈식별자_value)
	uc.작성자 = get_value(sheet, 작성자_value) 
	
	uc.유즈케이스명 = get_value(sheet, 유즈케이스명_value)
	uc.유즈케이스식별자 = get_value(sheet, 유즈케이스식별자_value)
	uc.작성일시 = get_value(sheet, 작성일시_value) 
	
	uc.관련액터 = get_value(sheet, 관련액터_value)
	uc.설명 = get_value(sheet, 설명_value)
	uc.사전조건 = get_value(sheet, 사전조건_value)
	uc.사후조건 = get_value(sheet, 사후조건_value)
	
	// 기본 흐름
	int i = -1
	for(row in 9..100) {
		String 번호 = get_value(sheet, "A${row}")
		
		if (번호 == "대안 흐름") {
			i = row + 2;
			break;
		}
		
		uc.기본흐름_단계설명.add(get_value(sheet, "B${row}"))
	}

	// 대안 흐름
	// 오류 처리 - 다시 하자 
	if (i < 0) 
		return uc
		
	for(row in i..100) {
		String 번호 = get_value(sheet, "A${row}")
		String 조건 = get_value(sheet, "B${row}")
		String 단계설명 = get_value(sheet, "C${row}")
		
		if (번호.length() <= 0) 
			break;
		
		uc.대안흐름_번호.add(get_value(sheet, "A${row}"))
		uc.대안흐름_조건.add(get_value(sheet, "B${row}"))
		uc.대안흐름_단계설명.add(get_value(sheet, "C${row}"))
	}
	
	return uc
}

//// SCRIPT !!!

def word = new ActiveXObject('Word.Application')
word.Documents.Add()
def doc = word.ActiveDocument
word.Visible = true


def usecase_file = new File('SK u-City 유즈케이스 템플릿_v1.0.xls').canonicalPath
println usecase_file

def xls = new ActiveXObject('Excel.Application')
def workbooks = xls.Workbooks
def workbook = workbooks.Open(usecase_file)

for (sheet in workbook.Sheets) {
	println "" + sheet.Name.value
	def uc = print_sheet(sheet)
	add_usecase_table(uc, word)
}
print "done"


workbook.Close()
xls.Quit()
//xls.release()

// doc.SaveAs(new File(".\\document.doc").canonicalPath)
// word.Quit()
println "done"