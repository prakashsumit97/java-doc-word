import org.codehaus.groovy.scriptom.ActiveXObject

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

UseCase uc = new UseCase()


def word = new ActiveXObject('Word.Application')
word.Documents.Add()
def doc = word.ActiveDocument
word.Visible = true

add_usecase_table(uc, word)
add_usecase_table(uc, word)

print "done"