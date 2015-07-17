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

	public ExcelFormat(File fileFormat)
			throws Exception
	{
		this(new FileInputStream(fileFormat));
	}

	public ExcelFormat(InputStream fileFormat) throws Exception
	{
		this.fileFormat = (FileFormatDefinition) this.xmlParser.fromXml(fileFormat, getAliasTypeMap(), getConverterList());
		checkDuplicateSheetConfig();
		setRevertReference();
	}

	public FileFormatDefinition getFileFormat()
	{
		return fileFormat;
	}

	private void checkDuplicateSheetConfig() throws Exception
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
					if (StringUtils.containsOnly(sheetIndexs, " 0123456789,"))
					{
						StringTokenizer st = new StringTokenizer(sheetIndexs, ",");
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
					else
					{
						throw new Exception("invalid character for attribute [sheetIndexs: '" + sheetIndexs + "']");
					}
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
