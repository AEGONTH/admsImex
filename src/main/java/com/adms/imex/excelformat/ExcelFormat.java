package com.adms.imex.excelformat;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import com.adms.imex.util.ObjectUtil;
import com.adms.imex.util.XmlParser;
import com.thoughtworks.xstream.converters.Converter;

public class ExcelFormat {

	private FileFormatDefinition fileFormat;
	private XmlParser xmlParser = new XmlParser();
	private String fileFormatXml;

	public ExcelFormat(File fileFormat)
			throws Exception
	{
		this(new FileInputStream(fileFormat));
	}

	public ExcelFormat(InputStream fileFormat) throws Exception
	{
		this.fileFormat = (FileFormatDefinition) this.xmlParser.fromXml(fileFormat, getAliasTypeMap(), getConverterList());
		doReplicateSheetConfig();

		this.fileFormatXml = this.xmlParser.toXml(this.fileFormat, getAliasTypeMap(), getConverterList());

		setRevertReference();
	}

	public String exportFileFormat()
	{
		return this.fileFormatXml;
	}

	public FileFormatDefinition getFileFormat()
	{
		return fileFormat;
	}

	private void doReplicateSheetConfig() throws Exception
	{
		if (CollectionUtils.isNotEmpty(this.fileFormat.getDataSetDefinition().getSheetDefinitionList()))
		{
			ArrayList<SheetDefinition> removeList = new ArrayList<SheetDefinition>();
			ArrayList<SheetDefinition> cloneList = new ArrayList<SheetDefinition>();

			for (SheetDefinition s : this.fileFormat.getDataSetDefinition().getSheetDefinitionList())
			{
				String sheetIndexs = s.getSheetIndexs();

				if (s.getSheetIndex() != null && StringUtils.isNotBlank(sheetIndexs))
				{
					throw new Exception("sheetIndex conflicted: " + s);
				}

				if (StringUtils.isNotBlank(sheetIndexs))
				{
					String sheetIndexs1 = separateString(sheetIndexs);

					StringTokenizer st = new StringTokenizer(sheetIndexs1, ",");
					while (st.hasMoreElements())
					{
						String sheetIndex = st.nextElement().toString().trim();
						if (StringUtils.isNotBlank(sheetIndex))
						{
							SheetDefinition cs = (SheetDefinition) ObjectUtil.deepClone(s);
							cs.setSheetIndex(Integer.parseInt(sheetIndex));

							cloneList.add(cs);
						}
					}
					removeList.add(s);
				}
			}

			if (CollectionUtils.isNotEmpty(removeList))
			{
				this.fileFormat.getDataSetDefinition().getSheetDefinitionList().removeAll(removeList);
			}

			if (CollectionUtils.isNotEmpty(cloneList))
			{
				this.fileFormat.getDataSetDefinition().getSheetDefinitionList().addAll(cloneList);
			}
		}
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

	private void setRevertReference()
	{
		for (SheetDefinition s : this.fileFormat.getDataSetDefinition().getSheetDefinitionList())
		{
			s.setFileFormatDefinition(this.fileFormat);

			if (CollectionUtils.isNotEmpty(s.getCellDefinitionList()))
			{
				for (CellDefinition c : s.getCellDefinitionList())
				{
					c.setSheetDefinition(s);
				}
			}

			if (CollectionUtils.isNotEmpty(s.getRecordDefinitionList()))
			{
				for (RecordDefinition r : s.getRecordDefinitionList())
				{
					r.setSheetDefinition(s);

					if (CollectionUtils.isNotEmpty(r.getCellDefinitionList()))
					{
						for (CellDefinition c : r.getCellDefinitionList())
						{
							c.setSheetDefinition(s);
						}
					}

					if (r.getBeginRecordCondition() != null)
					{
						r.getBeginRecordCondition().setSheetDefinition(s);
					}

					if (r.getEndRecordCondition() != null)
					{
						r.getEndRecordCondition().setSheetDefinition(s);
					}
				}
			}

			if (CollectionUtils.isNotEmpty(s.getColumnWidthDefinitionList()))
			{
				for (ColumnWidthDefinition c : s.getColumnWidthDefinitionList())
				{
					c.setSheetDefinition(s);
				}
			}

			if (CollectionUtils.isNotEmpty(s.getRowHeightDefinitionList()))
			{
				for (RowHeightDefinition r : s.getRowHeightDefinitionList())
				{
					r.setSheetDefinition(s);
				}
			}
		}
	}

	private Map<String, Class<?>> getAliasTypeMap()
	{
		Map<String, Class<?>> aliasTypeMap = new HashMap<String, Class<?>>();
		aliasTypeMap.put("FileFormat", FileFormatDefinition.class);
		aliasTypeMap.put("DataSet", DataSetDefinition.class);
		aliasTypeMap.put("DataSheet", SheetDefinition.class);
		aliasTypeMap.put("DataRecord", RecordDefinition.class);
		aliasTypeMap.put("DataCell", CellDefinition.class);
		aliasTypeMap.put("ColumnWidth", ColumnWidthDefinition.class);

		return aliasTypeMap;
	}

	private List<Converter> getConverterList()
	{
		List<Converter> converterList = new ArrayList<Converter>();

		return converterList;
	}

	private boolean isValidFileFormat()
			throws Exception
	{
		if (this.fileFormat == null)
		{
			throw new Exception("fileFormat is null, call setFileFormat() first!");
		}

		return true;
	}

	public void writeExcel(OutputStream output, DataHolder fileDataHolder)
			throws Exception
	{
		if (isValidFileFormat())
		{
			new ExcelFormatWriter(this, fileDataHolder).write(output);
		}
	}

	public DataHolder readExcel(InputStream input)
			throws Exception
	{
		if (!isValidFileFormat())
		{
			throw new Exception("invalid fileFormat");
		}

		return new ExcelFormatReader(this).read(input);
	}
}
