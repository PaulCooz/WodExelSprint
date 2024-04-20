#pragma once

namespace WodExelSprint {
	using namespace System;
	using namespace Microsoft::Office::Interop::Excel;

	ref class Sheet
	{
	private:
		Microsoft::Office::Interop::Excel::Application^ application;
		Workbook^ workbook;
		Worksheet^ worksheet;
		Range^ cells;

	public:
		Sheet(String^ path);
		String^ GetStr(int row, int clm);
	};
}
