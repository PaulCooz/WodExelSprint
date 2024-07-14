#include "Sheet.h"

namespace WodExelSprint {
	using namespace System::Text::RegularExpressions;

	Sheet::Sheet(String^ path)
	{
		this->application = gcnew ApplicationClass();
		this->workbook = this->application->Workbooks->Open(
			path,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing
		);
	}

	Worksheet^ Sheet::AddWorksheet(int indexFromBack)
	{
		auto sheets = this->workbook->Sheets;
		auto newSheet = sheets->Add(Type::Missing, sheets[sheets->Count - indexFromBack], Type::Missing, Type::Missing);
		return (Worksheet^)newSheet;
	}

	List<Worksheet^>^ Sheet::GetWorksheetsByName(String^ nameRegExp)
	{
		auto iterator = this->workbook->Worksheets->GetEnumerator();
		auto list = gcnew List<Worksheet^>();
		while (iterator->MoveNext())
		{
			auto worksheet = (Worksheet^)iterator->Current;
			if (Regex::IsMatch(worksheet->Name, nameRegExp))
			{
				list->Add(worksheet);
			}
		}
		return list;
	}

	String^ Sheet::GetStr(Worksheet^ worksheet, int row, int clm)
	{
		return ((Range^)worksheet->UsedRange->Cells[row, clm])->Text->ToString();
	}

	void Sheet::SetStr(Worksheet^ worksheet, int row, int clm, String^ str)
	{
		auto cells = (Range^)(worksheet->UsedRange->Cells[row, clm]);
		cells->Value2 = str;
	}

	String^ Sheet::GetStr(Worksheet^ worksheet, String^ range)
	{
		return ((Range^)worksheet->Range[range, Type::Missing])->Text->ToString();
	}

	void Sheet::SetStr(Worksheet^ worksheet, String^ range, String^ str)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->Merge(Type::Missing);
		cells->Value2 = str;
	}

	void Sheet::SetColor(Worksheet^ worksheet, String^ range, Object^ color)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->Merge(Type::Missing);
		cells->Interior->Color = color;
	}

	void Sheet::SetColWidth(Worksheet^ worksheet, String^ range, float value)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->ColumnWidth = value;
	}

	void Sheet::SetRowHeight(Worksheet^ worksheet, String^ range, float value)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->RowHeight = value;
	}

	void Sheet::SetFontBold(Worksheet^ worksheet, int row, int clm, bool value)
	{
		auto cells = (Range^)(worksheet->UsedRange->Cells[row, clm]);
		cells->Font->Bold = value;
	}

	void Sheet::SetFontBold(Worksheet^ worksheet, String^ range, bool value)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->Font->Bold = value;
	}

	void Sheet::SetHorAlign(Worksheet^ worksheet, int row, int clm, Object^ value)
	{
		auto cells = (Range^)(worksheet->UsedRange->Cells[row, clm]);
		cells->HorizontalAlignment = value;
	}

	void Sheet::SetHorAlign(Worksheet^ worksheet, String^ range, Object^ value)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->HorizontalAlignment = value;
	}

	void Sheet::SetVerAlign(Worksheet^ worksheet, int row, int clm, Object^ value)
	{
		auto cells = (Range^)(worksheet->UsedRange->Cells[row, clm]);
		cells->VerticalAlignment = value;
	}

	void Sheet::SetVerAlign(Worksheet^ worksheet, String^ range, Object^ value)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		cells->VerticalAlignment = value;
	}

	void Sheet::SetColor(Worksheet^ worksheet, int row, int clm, Object^ color)
	{
		((Range^)worksheet->UsedRange->Cells[row, clm])->Interior->Color = color;
	}

	void Sheet::SetNumberFormat(Worksheet^ worksheet, int row, int clm, String^ format)
	{
		((Range^)worksheet->UsedRange->Cells[row, clm])->NumberFormat = format;
	}

	void Sheet::SetBorder(Worksheet^ worksheet, int row, int clm, bool value)
	{
		auto cells = (Range^)(worksheet->UsedRange->Cells[row, clm]);
		if (value)
			cells->BorderAround(XlLineStyle::xlContinuous, XlBorderWeight::xlThin, XlColorIndex::xlColorIndexAutomatic, Type::Missing);
		else
			cells->Borders->LineStyle = XlLineStyle::xlLineStyleNone;
	}

	void Sheet::SetBorder(Worksheet^ worksheet, String^ range, bool value)
	{
		auto cells = (Range^)(worksheet->Range[range, Type::Missing]);
		if (value)
			cells->BorderAround(XlLineStyle::xlContinuous, XlBorderWeight::xlThin, XlColorIndex::xlColorIndexAutomatic, Type::Missing);
		else
			cells->Borders->LineStyle = XlLineStyle::xlLineStyleNone;
	}

	void Sheet::InsertRowUp(Worksheet^ worksheet, int row)
	{
		((Range^)worksheet->Rows[row, Type::Missing])->Insert(Type::Missing, Type::Missing);
	}

	void Sheet::InsertColLeft(Worksheet^ worksheet, int col)
	{
		((Range^)worksheet->Columns[col, Type::Missing])->Insert(XlInsertShiftDirection::xlShiftToRight, XlInsertFormatOrigin::xlFormatFromRightOrBelow);
	}

	void Sheet::SetVisible(bool value)
	{
		this->application->Visible = value;
	}
}
