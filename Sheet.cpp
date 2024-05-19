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

	void Sheet::SetColor(Worksheet^ worksheet, int row, int clm, XlRgbColor color)
	{
		((Range^)worksheet->UsedRange->Cells[row, clm])->Interior->Color = color;
	}

	void Sheet::SetVisible(bool value)
	{
		this->application->Visible = value;
	}
}