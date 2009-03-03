package com.sollink.docu.scripting

import org.apache.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import com.sollink.docu.dao.MemberDAO;
import com.sollink.docu.dao.CompoundDAO;
import com.sollink.docu.domain.Compound;
import com.sollink.docu.domain.Member;

import org.codehaus.groovy.scriptom.ActiveXObject

class XML2UML {	
	public void setCompoundDAO(CompoundDAO dao) {
		compoundDAO= dao;
	}
	
	public void gen_unit_testing_doc() {
//		대상  파일
		String input_file = "C:\\Documents and Settings\\sollink\\My Documents\\SKT\\미분류\\01. 산출물 생성 방안\\12월, 단위 테스트\\2SK u-City clover coverage check (version 1).csv"

//		대상 파일 분석 
		def file = new File(input_file)
		file.eachLine { line -> 
			String[] datum = line.split(",")
			println line
			println datum[0] + datum[1] + datum[2] + datum[3]
		}
		
//		신규 문서를 생성한다.
//		def word = new ActiveXObject('Word.Application')
//		word.Documents.Add()
//		def doc = word.ActiveDocument
//		word.Visible = true
//		
//		
//		def table = doc.Tables.Add(word.Selection.Range, 1, 3)
	}
	
	public ActiveXObject gen_word() {
		def word = new ActiveXObject('Word.Application')
		word.Documents.Add()
		def doc = word.ActiveDocument
		word.Visible = true
		
		return word
	}
	
	public void testing() {
		ActiveXObject word = gen_word()
		def doc = word.ActiveDocument
		def table = doc.Tables.Add(word.Selection.Range, 1, 3)
		
		TwoCellHeaderContents(table, 1, "클래스 명", "TBD", word)
		TwoCellHeaderContents(table, 2, "설명", "TBD", word)
		OneCellHeader(table, 3, "속성", word)
		ThreeCellHeader(table, 4, "이름", "타입", "설명", word)
		ThreeCellContents(table, 5, "TBD", "TBD", "TBD", word)
		OneCellHeader(table, 6, "메소드", word)
		ThreeCellHeaderContents(table, 7, "", "설명", "TBD", word)
		FourCellHeader(table, 8, "", "", "이름", "타입", word)
		FourCellContents(table, 9, "", "", "TBD", "TBD", word)
		FourCellHeader(table, 10, "", "", "이름", "타입", word)
		FourCellContents(table, 11, "", "", "TBD", "TBD", word)
		Merge(table, 10, 2, 11, 2, "반환값", word)
		Merge(table, 8, 2, 9, 2, "파라미터", word)
		Merge(table, 8, 1, 11, 1, "TBD", word)
	}
	
	public void Merge(def table, int row1, int col1, int row2, int col2, String msg, def word){
		table.Cell(row1, col1).Select();
		//		Unit:=wdLine (5), Count:=(row2 -row1), Extended:=wdExtend (1)
		if (row2-row1 > 0)
			word.Selection.MoveDown(5, row2-row1,1)
		//		Unit:=wdCharacter (1), Count:=(col2 -col1), Extended:=wdExtend (1)
		if( col2-col1 > 0)
			word.Selection.MoveRight(1, col2-col1, 1)
		word.Selection.Cells.Merge()
		table.Cell(row1, col1).Range.Text = msg
	}
	
	public void FourCellContents(def table, int row_idx, String content1, String content2, String content3, String content4, def word) {
		create_if_not_exist(row_idx, table, word)
		def row = table.Rows(row_idx)
		
		row.Cells(1).Split(1, 4)
		row.Cells(1).Width = 70
		row.Cells(2).Width = 70
		row.Cells(3).Width = 70
		row.Cells(4).Width = (300 - 70)
		row.Cells(1).Range.Text = content1
		row.Cells(2).Range.Text = content2
		row.Cells(3).Range.Text = content3
		row.Cells(4).Range.Text = content4
	}
	
	public void FourCellHeader(def table, int row_idx, String header1, String header2, String header3, String header4, def word) {
		FourCellContents(table, row_idx, header1, header2, header3, header4, word)
		
		def row = table.Rows(row_idx)
		row.Range.Shading.BackgroundPatternColor = gray_color 
	}
	
	public void ThreeCellContents(def table, int row_idx, String content1, String content2, String content3, def word) {
		create_if_not_exist(row_idx, table, word)
		def row = table.Rows(row_idx)
		
		row.Cells(1).Split(1, 3)
		row.Cells(1).Width = 70
		row.Cells(2).Width = 70
		row.Cells(3).Width = 300
		row.Cells(1).Range.Text = content1
		row.Cells(2).Range.Text = content2
		row.Cells(3).Range.Text = content3
	}
	
	public void ThreeCellHeader(def table, int row_idx, String header1, String header2, String header3, def word) {
		ThreeCellContents(table, row_idx, header1, header2, header3, word)
		
		def row = table.Rows(row_idx)
		row.Range.Shading.BackgroundPatternColor = gray_color 
	}
	
	public void ThreeCellHeaderContents(def table, int row_idx, String header1, String header2, String contents, def word) {
		ThreeCellHeader(table, row_idx, header1, header2, contents, word)
		
		def row = table.Rows(row_idx)
		row.Cells(3).Range.Shading.BackgroundPatternColor = white_color 
	}
	public void OneCellHeader(def table, int row_idx, String header, def word) {
		create_if_not_exist(row_idx, table, word)
		def row = table.Rows(row_idx)
		
		row.Cells(1).Range.Text = header
		row.Cells(1).Shading.BackgroundPatternColor = gray_color 
	}
	
	public void TwoCellHeaderContents(def table, int row_idx, String header, String msg, def word){
		create_if_not_exist(row_idx, table, word)
		def row = table.Rows(row_idx)

		row.Cells(1).Split(1, 2)
		row.Cells(1).Width = 70
		row.Cells(2).Width = 70 + 300
		row.Cells(1).Shading.BackgroundPatternColor = gray_color 
		row.Cells(1).Range.Text = header
		row.Cells(2).Range.Text = msg
	}
	int white_color = 16777215;
	int gray_color = 14737632;
	
	public void create_if_not_exist(int row_idx, def table, def word) {
		int last_idx = table.Rows.Last.Index;
		
		if (row_idx > last_idx) {
			for(int i=0; i<(row_idx - last_idx); i++) {
				table.Rows.Add()
			}
		}
		
		def row = table.Rows(row_idx)
		if (row.Cells.Count() > 1) {
			row.Select()
			word.Selection.Cells.Merge()
		}
		row.Range.Shading.BackgroundPatternColor = white_color //wdColorWhite
	}
	
	private static final Logger logger= Logger.getLogger(XML2UML.class)
	
	public boolean save_xml_file_to_db(String file_name, CompoundDAO compoundDAO, MemberDAO memberDAO) {
		def xml_file = new File(file_name).canonicalPath
		logger.debug("XML file name ${xml_file}")
		def doxygen = new XmlParser().parse(new File(xml_file))
		def compounddef = doxygen.compounddef[0]
		
//		Compound 처리
		Compound c = new Compound();		
//		1. id
		c.setId(compounddef.'@id');
//		2. kind
		c.setKind(compounddef.'@kind');
//		3. prot
		c.setProt(compounddef.'@prot');
//		4. name
		c.setName(compounddef.compoundname.text());
//		5. detaileddescription
		c.setDetaileddescription(toXML(compounddef.detaileddescription[0]));
//		6. basecompounddef
		c.setBasecompounddef(compounddef.basecompoundref.text());
		
//		존재할 경우, update 아니면 insert
		Compound exist_c = compoundDAO.selectByPrimaryKey(compounddef.'@id');
		if (exist_c == null) {
			compoundDAO.insert(c)
			logger.debug("Insert: ${compounddef.'@id'}")
		}
		else {
			compoundDAO.updateByPrimaryKey(c);
			logger.debug("Update: ${compounddef.'@id'}")
		}

//		Member 처리
		for(sectiondef in compounddef.sectiondef) {
			if (sectiondef.'@kind'.startsWith("public")) {
				for(memberdef in sectiondef.memberdef) {
					Member m = new Member();
					
//					1. id
					m.setId(memberdef.'@id');
//					2. prot
					m.setProt(memberdef.'@id');
//					3. kind
					m.setKind(memberdef.'@kind');
//					4. static_prefix
					m.setStaticPrefix(memberdef.'@static');
//					5. name
					m.setName(memberdef.name.text());
//					6. type_name
					m.setTypeName(memberdef.type.text());
//					7. detaileddescription
					m.setDetaileddescription(toXML(memberdef.detaileddescription[0]));
//					8. params
					m.setParams(memberdef.argsstring.text());					
//					9. compound_id
					m.setCompoundId(c.getId());

//					존재할 경우, update 아니면 insert
					Member exist_m = memberDAO.selectByPrimaryKey(memberdef.'@id');
					if (exist_m == null) {
						memberDAO.insert(m)
						logger.debug("Insert: ${m.getName()} - ${memberdef.'@id'}")
					}
					else {
						memberDAO.updateByPrimaryKey(m);
						logger.debug("Update: ${m.getName()} ${memberdef.'@id'}")
					}
				}
			}
		}

		return true
	}
	
	public boolean proc(String file_name, CompoundDAO compoundDAO) {
		def xml_file = new File(file_name).canonicalPath
		logger.debug("XML file name ${xml_file}")
		
//		Compound 처리

		def doxygen = new XmlParser().parse(new File(xml_file))
		def compounddef = doxygen.compounddef[0]
		
		Compound c = new Compound();
		c.setId(compounddef.'@id');
		c.setKind(compounddef.'@kind');
		c.setProt(compounddef.'@prot');
		
		logger.debug("Compound Creation ${c.getId()}")
		logger.debug("Compound name ${compounddef.compoundname.text()}")
		logger.debug("Compound detailed description ${toXML(compounddef.detaileddescription[0])}")
		logger.debug("Compound base compound def ${compounddef.basecompoundref.text()}")
		
		for(sectiondef in compounddef.sectiondef) {
			logger.debug("Sections ${sectiondef.'@kind'}")
			if (sectiondef.'@kind'.startsWith("public")) {
				for(memberdef in sectiondef.memberdef) {
					logger.debug("Member id ${memberdef.'@id'}")
					logger.debug("\t 1. Member prot ${memberdef.'@prot'}")
					logger.debug("\t 2. Member kind ${memberdef.'@kind'}")
					logger.debug("\t 3. Member static_prefix ${memberdef.'@static'}")
					logger.debug("\t 4. Member name ${memberdef.name.text()}")
					logger.debug("\t 5. Member type name ${memberdef.type.text()}")
					logger.debug("\t 6. Member detailed descrption ${toXML(memberdef.detaileddescription[0])}")
					logger.debug("\t 7. Member params ${memberdef.argsstring.text()}")
				}
			}
		}
		
		Compound exist = compoundDAO.selectByPrimaryKey(compounddef.'@id');
		if (exist == null) {
			compoundDAO.insert(c)
			logger.debug("Insert")
		}
		else {
			compoundDAO.updateByPrimaryKey(c);
			logger.debug("Update")
		}
		
		return true;
	}
	
	String file_name
	
	public XML2UML(){
	}
	
	public XML2UML(String file_name) {
		this.file_name = file_name
	}
	
	
	public String toXML(groovy.util.Node n) {
		if (n == null)
			return ""
		
		StringBuffer buf = new StringBuffer();
		buf.append("<${n.name()}")
		n.attributes().each {
			key, value -> buf.append(" ${key}=\"${value}\"")
		}
		buf.append("> ")
		n.children().each {
			node ->
			if (node instanceof String )
				buf.append(node)
			else
				buf.append(toXML(node))
		}
		
		buf.append("</${n.name()}>")
		return buf.toString();
	}
	
	public String png_file_name(String prefix) {
		String[] frag = file_name.split("\\.");
		String png_file_name = prefix + "\\" + frag[0] + "__coll__graph.png"
		return png_file_name
	}
	
	public void proc() {
		def xml_file = new File(file_name).canonicalPath
		println xml_file
		
		def doxygen = new XmlParser().parse(new File(xml_file))
		def compounddef = doxygen.compounddef[0]
		assert 'compounddef' == compounddef.name()
		println compounddef.name()
		println "id= ${compounddef.'@id'}"
		println "kind= ${compounddef.'@kind'}"
		println "prot= ${compounddef.'@prot'}"
//		println "briefdescription= ${compounddef.briefdescription.text()}"
		println "## to XML ${toXML(compounddef.detaileddescription[0])}"
		println "detaileddescription= ${compounddef.detaileddescription[0]}"
		
//		String xml_header = "<?xml version='1.0' encoding='UTF-8' standalone='no'?>"
		String xml_string = toXML(compounddef.detaileddescription[0])
		def xml_node = new XmlParser().parseText(xml_string)
		println "xml_node = ${xml_node}"
		
//		assert xml_node == compounddef.detaileddescription[0]
//		println "inbodydescription= ${compounddef.inbodydescription.text()}"
		
		for(sectiondef in compounddef.sectiondef) {
			println "\t sectiondef=${sectiondef.'@kind'}"
			String kind= sectiondef.'@kind'
			if (kind.startsWith("public")
			 || kind.startsWith("package")) {
				for(memberdef in sectiondef.memberdef ) {
//					TODO function v.s. variable 구분
//					TODO public/private 구분 / static 구분
//					TODO groovy.util.Node -> XML 함수를 작성하자...
					println "\t\t memberdef=${memberdef.name.text()}"
					println "\t\t memberdef=${memberdef.type.text()}"

//					println "\t\t briefdescription= ${memberdef.briefdescription}"
//					memberdef.detaileddescription[0]*.print(System.out);
					println "\t\t detaileddescription= ${memberdef.detaileddescription[0]*.value()}"
//					println "\t\t inbodydescription= ${memberdef.inbodydescription}"

					for(param in memberdef.param) {
						println "\t\t\t param=${param.declname.text()}:${param.type.text()}"
					}
				}
			}
		}
	}
	
	static void main(args) {
		XML2UML e = new XML2UML("classcom_1_1sk_1_1ucity_1_1dv_1_1admin_1_1insp_1_1web_1_1_insp_item_controller.xml");
//		String png_file_name = e.png_file_name("C:\\Documents and Settings\\sollink\\My Documents\\SKT\\doc\\html")
//		File f = new File(png_file_name)
//	
//		println f
//		println f.exists()
//		
//		File dir = new File("C:\\Documents and Settings\\sollink\\My Documents\\SKT\\doc\\html")
//		println dir.isDirectory()
//		dir.eachFile{
//			String name = it.name;
//			if (name.startsWith("classcom_1_1sk_1_1ucity_1_1dv_1_1admin_1_1insp_1_1web_1_1_insp_item_controller") 
//					&& name.endsWith(".png"))
//				println name
//		}
//		e.proc()
		e.gen_unit_testing_doc()
	}
}