package com.sollink.docu.dao;

import com.sollink.docu.domain.Compound;
import com.sollink.docu.domain.CompoundExample;
import java.util.List;

public interface CompoundDAO {

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	void insert(Compound record);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int updateByPrimaryKey(Compound record);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int updateByPrimaryKeySelective(Compound record);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	List<Compound> selectByExample(CompoundExample example);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	Compound selectByPrimaryKey(String id);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int deleteByExample(CompoundExample example);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int deleteByPrimaryKey(String id);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int countByExample(CompoundExample example);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int updateByExampleSelective(Compound record, CompoundExample example);

	/**
	 * This method was generated by Abator for iBATIS. This method corresponds to the database table COMPOUND
	 * @abatorgenerated  Fri Sep 19 17:02:23 KST 2008
	 */
	int updateByExample(Compound record, CompoundExample example);
}