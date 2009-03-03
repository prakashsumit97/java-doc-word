package com.sollink.docu.test;

import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.TestExecutionListeners;
import org.springframework.test.context.junit38.AbstractJUnit38SpringContextTests;
import org.springframework.test.context.transaction.TransactionConfiguration;
import org.springframework.test.context.transaction.TransactionalTestExecutionListener;
import org.springframework.transaction.annotation.Transactional;

import com.sollink.docu.scripting.XML2UML;

@ContextConfiguration(locations = {
		"file:./applicationContext.xml",
		"file:./applicationContext-ibatis.xml",})
public class ScriptTest extends AbstractJUnit38SpringContextTests {
	public void test_succ() {
		String file_name = "classcom_1_1sk_1_1ucity_1_1dv_1_1admin_1_1insp_1_1web_1_1_insp_item_controller.xml";
		
		XML2UML gen = new XML2UML("");
		gen.testing();
	}
}
