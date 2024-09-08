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
		Worksheet^ Sheet::AddWorksheet(int indexFromBack);
		List<Worksheet^>^ Sheet::GetWorksheetsByName(String^ nameRegExp);
		String^ Sheet::GetStr(Worksheet^ worksheet, int row, int clm);
		void Sheet::SetStr(Worksheet^ worksheet, int row, int clm, String^ str);
		String^ Sheet::GetStr(Worksheet^ worksheet, String^ range);
		void Sheet::SetStr(Worksheet^ worksheet, String^ range, String^ str);
		void Sheet::SetColor(Worksheet^ worksheet, int row, int clm, Object^ color);
		void Sheet::SetColor(Worksheet^ worksheet, String^ range, Object^ color);
		void Sheet::SetColWidth(Worksheet^ worksheet, String^ range, float value);
		void Sheet::SetRowHeight(Worksheet^ worksheet, String^ range, float value);
		void Sheet::SetFontBold(Worksheet^ worksheet, int row, int clm, bool value);
		void Sheet::SetFontBold(Worksheet^ worksheet, String^ range, bool value);
		void Sheet::SetHorAlign(Worksheet^ worksheet, int row, int clm, Object^ value);
		void Sheet::SetHorAlign(Worksheet^ worksheet, String^ range, Object^ value);
		void Sheet::SetVerAlign(Worksheet^ worksheet, int row, int clm, Object^ value);
		void Sheet::SetVerAlign(Worksheet^ worksheet, String^ range, Object^ value);
		void Sheet::SetNumberFormat(Worksheet^ worksheet, int row, int clm, String^ format);
		void Sheet::SetBorder(Worksheet^ worksheet, int row, int clm, bool value);
		void Sheet::SetBorder(Worksheet^ worksheet, String^ range, bool value);
		void Sheet::InsertRowUp(Worksheet^ worksheet, int row);
		void Sheet::InsertColLeft(Worksheet^ worksheet, int col);
		void SetVisible(bool value);

		String^ ColIntToStr(int col) {
			String^ s = "";
			while (col > 0) {
				s += (System::Char)('A' + ((col - 1) % ('Z' - 'A' + 1)));
				col /= ('Z' - 'A' + 1);
			}
			return s;
		}
	};
}
