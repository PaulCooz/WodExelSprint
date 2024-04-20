#include "Sheet.h"

namespace WodExelSprint {
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

		this->worksheet = (Worksheet^)(this->workbook->ActiveSheet);
		this->cells = this->worksheet->UsedRange->Cells;

		this->application->Visible = true;
	}

	String^ Sheet::GetStr(int row, int clm)
	{
		return ((Range^)cells[row, clm])->Text->ToString();
	}
}
