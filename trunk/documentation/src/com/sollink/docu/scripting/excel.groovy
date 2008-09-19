import org.codehaus.groovy.scriptom.ActiveXObject

class UseCase {
	public String ���θ���;
	public String ���θ��ĺ���;
	public String �ۼ���;
	public String �������̽���;
	public String �������̽��ĺ���
	public String �ۼ��Ͻ�;
	public String ���þ���;
	public String ����;
	public String ��������;
	public String ��������;
	
	public Vector<String> �⺻�帧_�ܰ輳�� = new Vector<String>();
	
	public Vector<String> ����帧_��ȣ = new Vector<String>();
	public Vector<String> ����帧_���� = new Vector<String>();
	public Vector<String> ����帧_�ܰ輳�� = new Vector<String>();
	
	def print() {
		println "Use Case"
		println ���θ���
		println ���θ��ĺ���
		println �ۼ���
		println �������̽���
		println �������̽��ĺ���
		println �ۼ��Ͻ�
		println ���þ���
		println ����
		println ��������
		println ��������
		println �⺻�帧_�ܰ輳��
		println ����帧_��ȣ
		println ����帧_����
		println ����帧_�ܰ輳��
	}
}

def add_usecase_table(uc, word) {
	def doc = word.ActiveDocument
	
	word.Selection.MoveDown(5, 3)
	word.Selection.TypeText("\n")
	def range = word.Selection.Range;
	
	//
	range.Select()
	int length = 8 + uc.�⺻�帧_�ܰ輳��.size() + 2 +  uc.����帧_�ܰ輳��.size() 
	def table = doc.Tables.Add(range,length,6)
	
	// 1��
	table.Cell(1,1).Range.Text = "���� ��� ��"
	table.Cell(1,1).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(1,2).Range.Text = uc.���θ���
	
	table.Cell(1,3).Range.Text = "���� ��� �ĺ���"
	table.Cell(1,3).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(1,4).Range.Text = uc.���θ��ĺ���
	
	table.Cell(1,5).Range.Text = "�ۼ���"
	table.Cell(1,5).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(1,6).Range.Text = uc.�ۼ���
	
	// 2��
	table.Cell(2,1).Range.Text = "�������̽� ��"
	table.Cell(2,1).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(2,2).Range.Text = uc.�������̽���
	
	table.Cell(2,3).Range.Text = "�������̽� �ĺ���"
	table.Cell(2,3).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(2,4).Range.Text = uc.�������̽��ĺ���
	
	table.Cell(2,5).Range.Text = "�ۼ� �Ͻ�"
	table.Cell(2,5).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(2,6).Range.Text = uc.�ۼ��Ͻ�
	
	// 3�� ~ 6��
	table.Cell(3,1).Range.Text = "���þ���"
	table.Cell(3,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(4,1).Range.Text = "����"
	table.Cell(4,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(5,1).Range.Text = "��������"
	table.Cell(5,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(6,1).Range.Text = "��������"
	table.Cell(6,1).Range.Shading.BackgroundPatternColor = 14737632
	
	table.Cell(3,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(3,2).Range.Text = uc.���þ���
	
	table.Cell(4,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(4,2).Range.Text = uc.����
	
	table.Cell(5,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(5,2).Range.Text = uc.��������
	
	table.Cell(6,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(6,2).Range.Text = uc.��������
	
	// �⺻ �帧 (7, 8)
	table.Cell(7,1).Select()
	word.Selection.MoveRight(1, 5, 1)
	word.Selection.Cells.Merge()
	table.Cell(7,1).Range.Text = "�⺻ �帧"
	table.Cell(7,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(7,1).Range.ParagraphFormat.Alignment = 1
	
	
	table.Cell(8,1).Range.Text = "��ȣ"
	table.Cell(8,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(8,1).Range.ParagraphFormat.Alignment = 1
	
	table.Cell(8,2).Select()
	word.Selection.MoveRight(1, 4, 1)
	word.Selection.Cells.Merge()
	table.Cell(8,2).Range.Text = "�ܰ輳��"
	table.Cell(8,2).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(8,2).Range.ParagraphFormat.Alignment = 1
	
	//
	int row = 9
	for(int i=0; i<uc.�⺻�帧_�ܰ輳��.size(); i++) {
		table.Cell(row, 1).Range.Text = "${i+1}"
		table.Cell(row,2).Select()
		word.Selection.MoveRight(1, 4, 1)
		word.Selection.Cells.Merge()
		table.Cell(row, 2).Range.Text = uc.�⺻�帧_�ܰ輳��.get(i)
		
		row ++
	}
	
	// ��� �帧
	
	table.Cell(row,1).Select()
	word.Selection.MoveRight(1, 5, 1)
	word.Selection.Cells.Merge()
	table.Cell(row,1).Range.Text = "��� �帧"
	table.Cell(row,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,1).Range.ParagraphFormat.Alignment = 1
	
	row ++
	
	table.Cell(row,1).Range.Text = "��ȣ"
	table.Cell(row,1).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,1).Range.ParagraphFormat.Alignment = 1
	
	table.Cell(row,2).Range.Text = "����"
	table.Cell(row,2).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,2).Range.ParagraphFormat.Alignment = 1
		
	table.Cell(row,3).Select()
	word.Selection.MoveRight(1, 3, 1)
	word.Selection.Cells.Merge()
	table.Cell(row,3).Range.Text = "�ܰ輳��"
	table.Cell(row,3).Range.Shading.BackgroundPatternColor = 14737632
	table.Cell(row,3).Range.ParagraphFormat.Alignment = 1
	
	row ++
	
	//
	for(int i=0; i<uc.����帧_�ܰ輳��.size(); i++) {
		table.Cell(row, 1).Range.Text = uc.����帧_��ȣ.get(i)
		table.Cell(row, 2).Range.Text = uc.����帧_����.get(i)
		
		table.Cell(row, 3).Select()
		word.Selection.MoveRight(1, 3, 1)
		word.Selection.Cells.Merge()
		table.Cell(row, 3).Range.Text = uc.����帧_�ܰ輳��.get(i)
		
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
	def ���θ���_value = "B1"
	def ���θ��ĺ���_value = "D1"
	def �ۼ���_value = "F1"
			
	def �������̽���_value = "B2"
	def �������̽��ĺ���_value = "D2"
	def �ۼ��Ͻ�_value = "F2"
	
	def ���þ���_value = "B3"
	def ����_value = "B4"
	def ��������_value = "B5"
	def ��������_value = "B6"
		
	UseCase uc = new UseCase()
	
	// header
	uc.���θ��� = get_value(sheet, ���θ���_value)
	uc.���θ��ĺ��� = get_value(sheet, ���θ��ĺ���_value)
	uc.�ۼ��� = get_value(sheet, �ۼ���_value) 
	
	uc.�������̽��� = get_value(sheet, �������̽���_value)
	uc.�������̽��ĺ��� = get_value(sheet, �������̽��ĺ���_value)
	uc.�ۼ��Ͻ� = get_value(sheet, �ۼ��Ͻ�_value) 
	
	uc.���þ��� = get_value(sheet, ���þ���_value)
	uc.���� = get_value(sheet, ����_value)
	uc.�������� = get_value(sheet, ��������_value)
	uc.�������� = get_value(sheet, ��������_value)
	
	// �⺻ �帧
	int i = -1
	for(row in 9..100) {
		String ��ȣ = get_value(sheet, "A${row}")
		
		if (��ȣ == "��� �帧") {
			i = row + 2;
			break;
		}
		
		uc.�⺻�帧_�ܰ輳��.add(get_value(sheet, "B${row}"))
	}

	// ��� �帧
	// ���� ó�� - �ٽ� ���� 
	if (i < 0) 
		return uc
		
	for(row in i..100) {
		String ��ȣ = get_value(sheet, "A${row}")
		String ���� = get_value(sheet, "B${row}")
		String �ܰ輳�� = get_value(sheet, "C${row}")
		
		if (��ȣ.length() <= 0) 
			break;
		
		uc.����帧_��ȣ.add(get_value(sheet, "A${row}"))
		uc.����帧_����.add(get_value(sheet, "B${row}"))
		uc.����帧_�ܰ輳��.add(get_value(sheet, "C${row}"))
	}
	
	return uc
}

//// SCRIPT !!!

def word = new ActiveXObject('Word.Application')
word.Documents.Add()
def doc = word.ActiveDocument
word.Visible = true


def usecase_file = new File('SK u-City �������̽� ���ø�_v1.0.xls').canonicalPath
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