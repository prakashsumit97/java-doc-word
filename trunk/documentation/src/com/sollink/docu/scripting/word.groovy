import org.codehaus.groovy.scriptom.ActiveXObject

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

UseCase uc = new UseCase()


def word = new ActiveXObject('Word.Application')
word.Documents.Add()
def doc = word.ActiveDocument
word.Visible = true

add_usecase_table(uc, word)
add_usecase_table(uc, word)

print "done"