package com.adms.imex.test;

import java.util.StringTokenizer;

import org.apache.commons.lang3.StringUtils;

public class TestString {

	public static void main(String[] args)
	{
		String sheetIndexs = "0,1,2";
		System.out.println(StringUtils.containsOnly(sheetIndexs, " 0123456789,"));
		System.out.println(StringUtils.containsOnly("", " 0123456789,"));
		
		StringTokenizer st = new StringTokenizer(sheetIndexs, ",");
		while (st.hasMoreElements())
		{
			System.out.println("[" + st.nextElement().toString().trim() + "]");
		}
		
	}

}
