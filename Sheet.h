#pragma once

namespace WodExelSprint {
	using namespace System;
	using namespace System::Collections::Generic;
	using namespace Microsoft::Office::Interop::Excel;

	ref class Sheet
	{
	private:
		Microsoft::Office::Interop::Excel::Application^ application;
		Workbook^ workbook;

	public:
		Sheet(String^ path);
		List<Worksheet^>^ Sheet::GetWorksheetsByName(String^ nameRegExp);
		String^ Sheet::GetStr(Worksheet^ worksheet, int row, int clm);
		void Sheet::SetStr(Worksheet^ worksheet, int row, int clm, String^ str);
		void Sheet::SetColor(Worksheet^ worksheet, int row, int clm, XlRgbColor color);
		void Sheet::SetNumberFormat(Worksheet^ worksheet, int row, int clm, String^ format);
		void SetVisible(bool value);
	};
}
