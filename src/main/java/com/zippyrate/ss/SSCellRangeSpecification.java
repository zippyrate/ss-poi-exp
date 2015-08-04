package com.zippyrate.ss;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

public class SSCellRangeSpecification
{
	private static final int START_COLUMN_INDEX_POSITION = 1;
	private static final int START_ROW_NUMBER_POSITION = 2;
	private static final int FINISH_COLUMN_INDEX_POSITION = 3;
	private static final int FINISH_ROW_NUMBER_POSITION = 4;
	private static final int NUMBER_OF_CAPTURING_GROUPS = 4;

	private final String sheetName;
	private final String rangeSpecification;
	private final int startRowNumber, startColumnIndex, finishRowNumber, finishColumnIndex;
	private static final Pattern pattern = Pattern.compile("([a-zA-Z]+)([1-9][0-9]*):(\\+|[a-zA-Z]+)(\\+|[1-9][0-9]*)");

	public SSCellRangeSpecification(String cellRangeSpecification) throws SSException
	{
		this(cellRangeSpecification.substring(0, cellRangeSpecification.indexOf('!')), cellRangeSpecification
				.substring(cellRangeSpecification.indexOf('!') + 1));
	}

	public SSCellRangeSpecification(String sheetName, String cellRangeSpecification) throws SSException
	{
		this.sheetName = sheetName;

		Matcher matcher = pattern.matcher(cellRangeSpecification);

		if (!matcher.find() || matcher.groupCount() != NUMBER_OF_CAPTURING_GROUPS)
			throw new SSException("Invalid cell range specification " + cellRangeSpecification);

		this.rangeSpecification = matcher.group(0);

		this.startColumnIndex = createColumnNumberFromSpecification(matcher.group(START_COLUMN_INDEX_POSITION));
		this.startRowNumber = createRowNumberFromSpecification(matcher.group(START_ROW_NUMBER_POSITION));

		this.finishColumnIndex = createColumnNumberFromSpecification(matcher.group(FINISH_COLUMN_INDEX_POSITION));
		this.finishRowNumber = createRowNumberFromSpecification(matcher.group(FINISH_ROW_NUMBER_POSITION));
	}

	public String getRangeSpecification()
	{
		return this.rangeSpecification;
	}

	public String getSheetName()
	{
		return this.sheetName;
	}

	public int getStartColumnIndex()
	{
		return this.startColumnIndex;
	}

	public int getStartRowNumber()
	{
		return this.startRowNumber;
	}

	public boolean hasOpenFinishRow()
	{
		return finishRowNumber == -1;
	}

	public boolean hasOpenFinishColumn()
	{
		return finishColumnIndex == -1;
	}

	public int getFinishColumnIndex()
	{
		return this.finishColumnIndex;
	}

	public int getFinishRowNumber()
	{
		return this.finishRowNumber;
	}

	public int getFinishRowNumber(Sheet sheet) throws SSException
	{
		return hasOpenFinishRow() ? sheet.getLastRowNum() : this.finishRowNumber;
	}

	public int getFinishColumnIndex(Row row) throws SSException
	{
		return hasOpenFinishColumn() ? (int)row.getLastCellNum() : this.finishColumnIndex;
	}

	public CellReference getStartCellReference()
	{
		return new CellReference(sheetName, startRowNumber, startColumnIndex, false, false);
	}

	private static int createColumnNumberFromSpecification(String columnSpec) throws SSException
	{
		if (columnSpec.equals("+"))
			return -1;
		else {
			int columnIndex = CellReference.convertColStringToIndex(columnSpec);

			if (columnIndex == Integer.MAX_VALUE)
				throw new SSException("Column " + columnSpec + " out of range!");
			else
				return columnIndex;
		}
	}

	private static int createRowNumberFromSpecification(String rowSpec) throws SSException
	{
		if (rowSpec.equals("+"))
			return -1;
		else {
			int rowNumber = convertRowNumberString2RowNumber(rowSpec) - 1;

			if (rowNumber == Integer.MAX_VALUE)
				throw new SSException("Row " + rowSpec + " out of range!");
			else
				return rowNumber;
		}
	}

	public static int convertRowNumberString2RowNumber(String rowNumber) throws SSException
	{
		try {
			return Integer.parseInt(rowNumber);
		} catch (NumberFormatException e) {
			throw new SSException("Invalid row number " + rowNumber + ": " + e.getMessage(), e);
		}
	}

	@Override
	public String toString()
	{
		return sheetName + '!' + rangeSpecification;
	}
}
