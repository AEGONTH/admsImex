package com.adms.imex.test;

import java.util.StringTokenizer;

import org.apache.commons.lang3.StringUtils;

public class TestString {

	public static void main(String[] args) throws Exception
	{
		System.out.println(separateString("1, 3 - 5, 8 - 10"));
	}

	private static String separateString(String s) throws Exception
	{
		if (StringUtils.isNotBlank(s))
		{
			if (StringUtils.containsOnly(s, " -,1234567890"))
			{
				String result = "";
				String ss = s.replaceAll(" ", "");
				
				StringTokenizer st = new StringTokenizer(ss, ",");
				while (st.hasMoreElements())
				{
					String sss = st.nextElement().toString().trim();
					if (StringUtils.isNotBlank(sss))
					{
						result += ", " + translateRange(sss);
					}
				}

				return result.replaceFirst(", ", "");
			}
			else
			{
				throw new Exception("invalid character for attribute sheetIndexs[" + s + "]");
			}
		}
		else
		{
			return "";
		}
	}

	private static String translateRange(String s)
			throws Exception
	{
		if (StringUtils.isNotBlank(s))
		{
			String ss = s.replaceAll(" ", "");
			
			if (ss.indexOf("-") == -1)
			{
				return ss;
			}
			else
			{
				if (ss.indexOf(",-") != -1 || ss.indexOf("-,") != -1 || ss.startsWith("-") || ss.endsWith("-"))
				{
					throw new Exception("invalid range expression [" + s + "]");
				}

				String begin = ss.substring(0, ss.indexOf('-')).trim();
				String end = ss.substring(ss.indexOf('-') + 1).trim();
				return a(Integer.parseInt(begin), Integer.parseInt(end));
			}
		}
		else
		{
			return "";
		}
	}

	private static String a(int begin, int end) throws Exception
	{
		StringBuilder sb = new StringBuilder("");
		if (begin < end)
		{
			for (int i = begin; i <= end; i++)
			{
				sb.append(String.valueOf(i));

				if (i < end)
					sb.append(String.valueOf(", "));
			}
		}
		else
		{
			throw new Exception("invalid range expression begin number [" + begin + "] cannot more than end number [" + end + "]");
		}

		return sb.toString();
	}

}
