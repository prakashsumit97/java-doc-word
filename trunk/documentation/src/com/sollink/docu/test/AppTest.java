package com.sollink.docu.test;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.TestExecutionListeners;
import org.springframework.test.context.junit38.AbstractJUnit38SpringContextTests;
import org.springframework.test.context.transaction.TransactionConfiguration;
import org.springframework.test.context.transaction.TransactionalTestExecutionListener;
import org.springframework.transaction.annotation.Transactional;

import com.sollink.docu.dao.MemberDAO;
import com.sollink.docu.domain.Member;
import com.sollink.docu.domain.MemberExample;

@ContextConfiguration(locations = {
		"file:./applicationContext.xml",
		"file:./applicationContext-ibatis.xml",})
@TransactionConfiguration(transactionManager = "txManager", defaultRollback = false)
@Transactional
@TestExecutionListeners( { TransactionalTestExecutionListener.class })

public class AppTest extends AbstractJUnit38SpringContextTests{
	
	@Autowired
	private MemberDAO mdao;
	
	public void test_first() {
		Member m = new Member();
		m.setId("babo");
	
		mdao.insert(m);
		MemberExample e = new MemberExample();
		e.createCriteria().andIdEqualTo("babo");
		
		Member m2 = mdao.selectByExample(e).get(0);
		assertNotNull(m2);
	}
	
	public void test_succ() {
		assertTrue(true);
	}
}
