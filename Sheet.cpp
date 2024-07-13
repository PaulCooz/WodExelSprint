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

	Worksheet^ Sheet::AddWorksheet()
	{
		auto sheets = this->workbook->Sheets;
		auto newSheet = sheets->Add(Type::Missing, sheets[sheets->Count], Type::Missing, Type::Missing);
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

	void Sheet::SetColor(Worksheet^ worksheet, int row, int clm, Object^ color)
	{
		((Range^)worksheet->UsedRange->Cells[row, clm])->Interior->Color = color;
	}

	void Sheet::SetNumberFormat(Worksheet^ worksheet, int row, int clm, String^ format)
	{
		((Range^)worksheet->UsedRange->Cells[row, clm])->NumberFormat = format;
	}

	void Sheet::SetVisible(bool value)
	{
		this->application->Visible = value;
	}
}
