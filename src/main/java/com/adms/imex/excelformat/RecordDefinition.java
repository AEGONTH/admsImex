package com.adms.imex.excelformat;

import com.thoughtworks.xstream.annotations.XStreamAlias;
import com.thoughtworks.xstream.annotations.XStreamAsAttribute;

public class RecordDefinition extends CellDefinitionContainer {

	@XStreamAsAttribute
	private String listSourceName;

	@XStreamAsAttribute
	private Integer beginRow;

	@XStreamAsAttribute
	private Integer endRow;

	@XStreamAsAttribute
	private Boolean skipWhenNull;

	@XStreamAsAttribute
	private Integer offsetRow;

	@XStreamAsAttribute
	private String offsetFrom;

	@XStreamAlias("BeginRecordCondition")
	private BeginRecordCondition beginRecordCondition;

	@XStreamAlias("EndRecordCondition")
	private EndRecordCondition endRecordCondition;

	private SheetDefinition sheetDefinition;

	public String getListSourceName()
	{
		return listSourceName;
	}

	public void setListSourceName(String listSourceName)
	{
		this.listSourceName = listSourceName;
	}

	public Integer getBeginRow()
	{
		return beginRow;
	}

	public void setBeginRow(Integer beginRow)
	{
		this.beginRow = beginRow;
	}

	public Integer getEndRow()
	{
		return endRow;
	}

	public void setEndRow(Integer endRow)
	{
		this.endRow = endRow;
	}

	public Boolean getSkipWhenNull()
	{
		return skipWhenNull;
	}

	public void setSkipWhenNull(Boolean skipWhenNull)
	{
		this.skipWhenNull = skipWhenNull;
	}

	public Integer getOffsetRow()
	{
		return offsetRow;
	}

	public void setOffsetRow(Integer offsetRow)
	{
		this.offsetRow = offsetRow;
	}

	public String getOffsetFrom()
	{
		return offsetFrom;
	}

	public void setOffsetFrom(String offsetFrom)
	{
		this.offsetFrom = offsetFrom;
	}

	public BeginRecordCondition getBeginRecordCondition()
	{
		return beginRecordCondition;
	}

	public void setBeginRecordCondition(BeginRecordCondition beginRecordCondition)
	{
		this.beginRecordCondition = beginRecordCondition;
	}

	public EndRecordCondition getEndRecordCondition()
	{
		return endRecordCondition;
	}

	public void setEndRecordCondition(EndRecordCondition endRecordCondition)
	{
		this.endRecordCondition = endRecordCondition;
	}

	public SheetDefinition getSheetDefinition()
	{
		return sheetDefinition;
	}

	public void setSheetDefinition(SheetDefinition sheetDefinition)
	{
		this.sheetDefinition = sheetDefinition;

		if (this.beginRecordCondition != null)
		{
			this.beginRecordCondition.setSheetDefinition(sheetDefinition);
		}

		if (this.endRecordCondition != null)
		{
			this.endRecordCondition.setSheetDefinition(sheetDefinition);
		}
	}

	@Override
	public String toString()
	{
		return "RecordDefinition [listSourceName=" + listSourceName + ", beginRow=" + beginRow + ", endRow=" + endRow + ", skipWhenNull=" + skipWhenNull + ", offsetRow=" + offsetRow + ", offsetFrom=" + offsetFrom + ", beginRecordCondition=" + beginRecordCondition + ", endRecordCondition=" + endRecordCondition + "]";
	}

}
