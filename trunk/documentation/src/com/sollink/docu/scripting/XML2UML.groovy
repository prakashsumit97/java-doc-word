package com.sollink.docu.scripting

class XML2UML {
String file_name
	
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
		e.proc()
	}
}