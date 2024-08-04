#pragma once

#include "Sheet.h"
#include <algorithm>

namespace WodExelSprint {
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;
	using namespace System::Diagnostics;
	using namespace System::Text::RegularExpressions;
	using namespace System::Windows::Forms;
	using namespace Microsoft::Office::Interop;

	/// <summary>
	/// Summary for Main
	/// </summary>
	public ref class Main : public System::Windows::Forms::Form
	{
	private: System::Windows::Forms::Button^ ClearXlsxButton;

	private: System::Windows::Forms::TableLayoutPanel^ TableLayoutPanel;
	private: System::Windows::Forms::Button^ AddTeamXlsxButton;

	private: System::Windows::Forms::Button^ ValidateXlsxButton;
	private:


	public:
		Main(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~Main()
		{
		}
	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->ValidateXlsxButton = (gcnew System::Windows::Forms::Button());
			this->ClearXlsxButton = (gcnew System::Windows::Forms::Button());
			this->TableLayoutPanel = (gcnew System::Windows::Forms::TableLayoutPanel());
			this->AddTeamXlsxButton = (gcnew System::Windows::Forms::Button());
			this->TableLayoutPanel->SuspendLayout();
			this->SuspendLayout();
			// 
			// ValidateXlsxButton
			// 
			this->ValidateXlsxButton->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->ValidateXlsxButton->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->ValidateXlsxButton->Location = System::Drawing::Point(3, 3);
			this->ValidateXlsxButton->Name = L"ValidateXlsxButton";
			this->ValidateXlsxButton->Size = System::Drawing::Size(276, 66);
			this->ValidateXlsxButton->TabIndex = 0;
			this->ValidateXlsxButton->Text = L"validate xlsx";
			this->ValidateXlsxButton->UseVisualStyleBackColor = true;
			this->ValidateXlsxButton->Click += gcnew System::EventHandler(this, &Main::OpenXlsxFile);
			// 
			// ClearXlsxButton
			// 
			this->ClearXlsxButton->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->ClearXlsxButton->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->ClearXlsxButton->Location = System::Drawing::Point(3, 75);
			this->ClearXlsxButton->Name = L"ClearXlsxButton";
			this->ClearXlsxButton->Size = System::Drawing::Size(276, 66);
			this->ClearXlsxButton->TabIndex = 1;
			this->ClearXlsxButton->Text = L"clear xlsx";
			this->ClearXlsxButton->UseVisualStyleBackColor = true;
			this->ClearXlsxButton->Click += gcnew System::EventHandler(this, &Main::ClearXlsxButton_Click);
			// 
			// TableLayoutPanel
			// 
			this->TableLayoutPanel->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->TableLayoutPanel->AutoSize = true;
			this->TableLayoutPanel->ColumnCount = 1;
			this->TableLayoutPanel->ColumnStyles->Add((gcnew System::Windows::Forms::ColumnStyle(System::Windows::Forms::SizeType::Percent,
				100)));
			this->TableLayoutPanel->Controls->Add(this->ValidateXlsxButton, 0, 0);
			this->TableLayoutPanel->Controls->Add(this->ClearXlsxButton, 0, 1);
			this->TableLayoutPanel->Controls->Add(this->AddTeamXlsxButton, 0, 2);
			this->TableLayoutPanel->GrowStyle = System::Windows::Forms::TableLayoutPanelGrowStyle::FixedSize;
			this->TableLayoutPanel->Location = System::Drawing::Point(12, 12);
			this->TableLayoutPanel->Name = L"TableLayoutPanel";
			this->TableLayoutPanel->RowCount = 3;
			this->TableLayoutPanel->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
			this->TableLayoutPanel->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
			this->TableLayoutPanel->RowStyles->Add((gcnew System::Windows::Forms::RowStyle(System::Windows::Forms::SizeType::Percent, 33.33333F)));
			this->TableLayoutPanel->Size = System::Drawing::Size(282, 218);
			this->TableLayoutPanel->TabIndex = 2;
			// 
			// AddTeamXlsxButton
			// 
			this->AddTeamXlsxButton->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->AddTeamXlsxButton->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->AddTeamXlsxButton->Location = System::Drawing::Point(3, 147);
			this->AddTeamXlsxButton->Name = L"AddTeamXlsxButton";
			this->AddTeamXlsxButton->Size = System::Drawing::Size(276, 68);
			this->AddTeamXlsxButton->TabIndex = 2;
			this->AddTeamXlsxButton->Text = L"add team xlsx";
			this->AddTeamXlsxButton->UseVisualStyleBackColor = true;
			this->AddTeamXlsxButton->Click += gcnew System::EventHandler(this, &Main::AddTeamXlsxButton_Click);
			// 
			// Main
			// 
			this->ClientSize = System::Drawing::Size(306, 242);
			this->Controls->Add(this->TableLayoutPanel);
			this->Name = L"Main";
			this->TableLayoutPanel->ResumeLayout(false);
			this->ResumeLayout(false);
			this->PerformLayout();

		}

		Form^ prompt = nullptr;

		String^ ShowInputDialog(String^ text, String^ caption) {
			if (prompt != nullptr) {
				return "error";
			}

			prompt = gcnew Form();
			prompt->Width = 500;
			prompt->Height = 150;
			prompt->FormBorderStyle = System::Windows::Forms::FormBorderStyle::FixedDialog;
			prompt->Text = caption;
			prompt->StartPosition = FormStartPosition::CenterScreen;

			auto textLabel = gcnew System::Windows::Forms::Label();
			textLabel->Left = 50;
			textLabel->Top = 20;
			textLabel->Width = 400;
			textLabel->Text = text;

			auto textBox = gcnew System::Windows::Forms::TextBox();
			textBox->Left = 50;
			textBox->Top = 50;
			textBox->Width = 400;

			auto confirmation = gcnew System::Windows::Forms::Button();
			confirmation->Text = "Ok";
			confirmation->Left = 350;
			confirmation->Width = 100;
			confirmation->Top = 70;
			confirmation->DialogResult = System::Windows::Forms::DialogResult::OK;
			confirmation->Click += gcnew System::EventHandler(this, &Main::ConfirmInputDialog);

			prompt->Controls->Add(textBox);
			prompt->Controls->Add(confirmation);
			prompt->Controls->Add(textLabel);
			prompt->AcceptButton = confirmation;

			auto success = prompt->ShowDialog() == System::Windows::Forms::DialogResult::OK;
			return success ? textBox->Text : "";
		}

		void ConfirmInputDialog(System::Object^ sender, System::EventArgs^ e) {
			prompt->Close();
			prompt = nullptr;
		}

	private: System::Void OpenXlsxFile(System::Object^ sender, System::EventArgs^ e) {
		OpenFileDialog^ openFileDialog = gcnew OpenFileDialog;

		openFileDialog->InitialDirectory = ".";
		openFileDialog->Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
		openFileDialog->FilterIndex = 1;
		openFileDialog->RestoreDirectory = true;

		if (openFileDialog->ShowDialog() != System::Windows::Forms::DialogResult::OK)
			return;

		auto sheet = gcnew Sheet(openFileDialog->FileName);
		auto worksheet = sheet->GetWorksheetsByName("Sprint")[0];

		const static int VelocityCol = 3;
		const static int AbsenceCol = 4;
		const static int WaringCol = 6;

		const static int symLen = 3;
		static unsigned char symBytes[symLen] = { 0xE2, 0x9A, 0xA0 };
		auto symbol = System::Text::Encoding::UTF8->GetString(symBytes, symLen);

		auto inTeamRange = false;
		auto skipFirstPerson = false;
		for (auto r = 1; ; r++)
		{
			auto colB = sheet->GetStr(worksheet, r, VelocityCol);
			if (colB == "Velocity")
			{
				skipFirstPerson = true;
				if (!inTeamRange)
					inTeamRange = true;
				continue;
			}
			if (inTeamRange && String::IsNullOrEmpty(colB))
			{
				inTeamRange = false;
				break;
			}
			auto isNumber = System::Text::RegularExpressions::Regex::IsMatch(colB, "\\d+[\\.,]?\\d*");
			if (!isNumber || !inTeamRange || skipFirstPerson)
			{
				skipFirstPerson = false;
				continue;
			}

			auto value = System::Single::Parse(sheet->GetStr(worksheet, r, VelocityCol));
			auto notFull = value < 9;
			auto noReason = String::IsNullOrEmpty(sheet->GetStr(worksheet, r, AbsenceCol));
			if (notFull && noReason)
			{
				sheet->SetStr(worksheet, r, WaringCol, symbol);
			}
		}

		sheet->SetVisible(true);
	}

	private: System::Void ClearXlsxButton_Click(System::Object^ sender, System::EventArgs^ e) {
		OpenFileDialog^ openFileDialog = gcnew OpenFileDialog;

		openFileDialog->InitialDirectory = ".";
		openFileDialog->Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
		openFileDialog->FilterIndex = 1;
		openFileDialog->RestoreDirectory = true;

		if (openFileDialog->ShowDialog() != System::Windows::Forms::DialogResult::OK)
			return;

		auto sheet = gcnew Sheet(openFileDialog->FileName);
		auto worksheet = sheet->GetWorksheetsByName("Sprint")[0];
		String^ TeamNameTemp = "(.*team)|(DevOps)";
		auto culture = System::Globalization::CultureInfo::InvariantCulture;

		{
			const static int focusFactorCol = 1;
			int focusFactorRow = 1;
			while (String::Compare(sheet->GetStr(worksheet, focusFactorRow, focusFactorCol), "Actual focus factor", StringComparison::InvariantCultureIgnoreCase) != 0) {
				focusFactorRow++;
			}

			auto dict = gcnew System::Collections::Generic::Dictionary<String^, float>();
			int summaryRow = focusFactorRow - 5;
			int summaryCol = 2;
			while (!String::IsNullOrEmpty(sheet->GetStr(worksheet, summaryRow, summaryCol))) {
				auto name = sheet->GetStr(worksheet, summaryRow, summaryCol);
				auto isTeam = Regex::IsMatch(name, TeamNameTemp);
				if (isTeam) {
					auto value = System::Single::Parse(sheet->GetStr(worksheet, focusFactorRow, summaryCol), culture);
					dict->Add(name, value);
				}
				summaryCol++;
			}

			worksheet = sheet->GetWorksheetsByName("Focus factor")[0];
			auto lastRow = 1;
			while (!String::IsNullOrEmpty(sheet->GetStr(worksheet, lastRow, 1)))
				lastRow++;
			lastRow--;

			auto lastSprint = sheet->GetStr(worksheet, lastRow - 1, 1);
			auto sprint = System::Int32::Parse(lastSprint) + 1;
			sheet->SetStr(worksheet, lastRow + 1, 1, sheet->GetStr(worksheet, lastRow, 1));
			sheet->SetStr(worksheet, lastRow, 1, sprint.ToString());
			auto col = 2;
			while (!String::IsNullOrEmpty(sheet->GetStr(worksheet, 1, col))) {
				auto team = sheet->GetStr(worksheet, 1, col);
				if (dict->ContainsKey(team) && dict[team] > 0.0) {
					sheet->SetNumberFormat(worksheet, lastRow, col, "###%");
					sheet->SetStr(worksheet, lastRow, col, dict[team].ToString(culture));
				}
				else {
					sheet->SetStr(worksheet, lastRow, col, "-");
				}

				auto sum = 0.0;
				auto count = 0.0;
				auto startRow = std::max(2, lastRow - 9);

				for (auto i = startRow; i <= lastRow; i++) {
					auto m = Regex::Match(sheet->GetStr(worksheet, i, col), "\\d+");
					if (m->Success) {
						sum += Single::Parse(m->Value) / 100.0;
						count += 1.0;
					}
				}

				auto avg = sum / count;
				sheet->SetNumberFormat(worksheet, lastRow + 1, col, "###%");
				sheet->SetStr(worksheet, lastRow + 1, col, avg.ToString(culture));

				col++;
			}
		}

		const static int VelocityCol = 3;
		const static int AbsenceCol = 4;
		const static int WaringCol = 6;

		const static int symLen = 3;
		static unsigned char symBytes[symLen] = { 0xE2, 0x9A, 0xA0 };
		auto symbol = System::Text::Encoding::UTF8->GetString(symBytes, symLen);

		worksheet = sheet->GetWorksheetsByName("Sprint")[0];
		auto inTeamRange = false;
		auto skipFirstPerson = false;
		for (auto r = 1; ; r++)
		{
			auto colB = sheet->GetStr(worksheet, r, VelocityCol);
			if (colB == "Velocity")
			{
				skipFirstPerson = true;
				if (!inTeamRange)
					inTeamRange = true;
				continue;
			}
			if (inTeamRange && String::IsNullOrEmpty(colB))
			{
				inTeamRange = false;
				break;
			}
			auto isNumber = System::Text::RegularExpressions::Regex::IsMatch(colB, "-?\\d+[\\.,]?\\d*");
			if (!isNumber || !inTeamRange)
			{
				continue;
			}

			sheet->SetStr(worksheet, r, VelocityCol, skipFirstPerson ? "0" : "9");
			sheet->SetStr(worksheet, r, AbsenceCol, "");
			sheet->SetStr(worksheet, r, WaringCol, "");

			skipFirstPerson = false;
		}

		{
			const int dateRow = 1;
			const int dateStartCol = 2;
			const int dateFinishCol = 3;
			sheet->SetStr(worksheet, dateRow, dateStartCol, "");
			sheet->SetStr(worksheet, dateRow, dateFinishCol, "");
		}

		auto worksheets = sheet->GetWorksheetsByName(TeamNameTemp);
		for each (worksheet in worksheets)
		{
			auto name = worksheet->Name;
			auto checkForEmptyCol = 1;
			auto numbersVisited = 0;
			while (String::Compare(sheet->GetStr(worksheet, 2, checkForEmptyCol), "Unused velocity:") != 0)
				checkForEmptyCol++;
			checkForEmptyCol++;

			auto topRow = 4;
			while (!String::IsNullOrEmpty(sheet->GetStr(worksheet, topRow, checkForEmptyCol))) {
				sheet->SetStr(worksheet, topRow, 1, "");
				sheet->SetStr(worksheet, topRow, 4, "");
				sheet->SetStr(worksheet, topRow + 1, 4, "");
				for (int j = 5; j <= checkForEmptyCol; j++)
					sheet->SetStr(worksheet, topRow + 1, j, "");

				topRow += 3;
			}
		}

		sheet->SetVisible(true);
	}

	private: String^ ColIntToStr(int col) {
		String^ s = "";
		while (col > 0) {
			s += (System::Char)('A' + ((col - 1) % ('Z' - 'A' + 1)));
			col /= ('Z' - 'A' + 1);
		}
		return s;
	}

	private: System::Void AddTeamXlsxButton_Click(System::Object^ sender, System::EventArgs^ e) {
		OpenFileDialog^ openFileDialog = gcnew OpenFileDialog;

		openFileDialog->InitialDirectory = ".";
		openFileDialog->Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
		openFileDialog->FilterIndex = 1;
		openFileDialog->RestoreDirectory = true;

		if (openFileDialog->ShowDialog() != System::Windows::Forms::DialogResult::OK)
			return;

		auto sheet = gcnew Sheet(openFileDialog->FileName);
		auto newWorksheet = sheet->AddWorksheet(2);
		newWorksheet->Name = ShowInputDialog("enter team name:", "input");

		List<String^>^ developers = gcnew List<String^>();
		auto countQA = 0;
		{
			int countDevelopers = Int32::Parse(ShowInputDialog("enter count or developers:", "input"));

			for (int i = 0; i < countDevelopers; i++)
				developers->Add("");

			int front = 0;
			int back = developers->Count - 1;
			for (int i = 0; i < developers->Count; i++) {
				auto name = ShowInputDialog("enter developer #" + (i + 1) + " name:", "input");
				if (name->StartsWith("[QA]"))
				{
					developers[back--] = name;
					countQA++;
				}
				else
				{
					developers[front++] = name;
				}
			}
		}
		auto lastColLetter = ColIntToStr(5 + developers->Count - 1);

		sheet->SetStr(newWorksheet, "A1:J1", newWorksheet->Name + " planning table");
		sheet->SetFontBold(newWorksheet, "A1:J1", true);
		sheet->SetBorder(newWorksheet, "A1:J1", true);
		sheet->SetStr(newWorksheet, "A2:A3", "User stories");
		sheet->SetFontBold(newWorksheet, "A2:A3", true);
		sheet->SetBorder(newWorksheet, "A2:A3", true);
		sheet->SetStr(newWorksheet, "B2:C2", "Estimated SP:");
		sheet->SetStr(newWorksheet, 2, 4, "=SUMIFS('" + newWorksheet->Name + "'!$D$4:$D$102, '" + newWorksheet->Name + "'!$B$4:$B$102, \"Dev\")");
		sheet->SetFontBold(newWorksheet, "B2:D2", true);
		sheet->SetBorder(newWorksheet, "B2:D2", true);
		sheet->SetStr(newWorksheet, "E2:F2", "Actual focus factor:");
		sheet->SetFontBold(newWorksheet, "E2:G2", true);
		sheet->SetBorder(newWorksheet, "E2:G2", true);
		sheet->SetStr(newWorksheet, "H2:I2", "Unused velocity:");
		sheet->SetFontBold(newWorksheet, "H2:J2", true);
		sheet->SetBorder(newWorksheet, "H2:J2", true);
		sheet->SetStr(newWorksheet, "B3:D3", "Total estimation");
		sheet->SetFontBold(newWorksheet, "B3:D3", true);
		sheet->SetBorder(newWorksheet, "B3:D3", true);
		sheet->SetStr(newWorksheet, "E3:J3", "Teammates estimations");
		sheet->SetFontBold(newWorksheet, "E3:J3", true);
		sheet->SetBorder(newWorksheet, "E3:J3", true);

		auto colorDialog = gcnew ColorDialog();
		System::Drawing::Color color;
		if (colorDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
			color = colorDialog->Color;
		auto colorQA = System::Drawing::Color::FromArgb(255, 217, 102);

		sheet->SetColWidth(newWorksheet, "A:A", 63);
		sheet->SetRowHeight(newWorksheet, "A:A", 22);
		sheet->SetFontBold(newWorksheet, 1, 1, true);

		sheet->SetHorAlign(newWorksheet, "A1:" + lastColLetter + "12", XlHAlign::xlHAlignCenter);
		sheet->SetVerAlign(newWorksheet, "A1:" + lastColLetter + "12", XlHAlign::xlHAlignCenter);
		sheet->SetHorAlign(newWorksheet, "B4:D102", XlHAlign::xlHAlignRight);

		for (int row = 4; row <= 10; row += 3) {
			for (int i = 0; i < developers->Count; i++) {
				sheet->SetStr(newWorksheet, row, 5 + i, developers[i]);
				if (developers[i]->StartsWith("[QA]"))
					sheet->SetColor(newWorksheet, row, 5 + i, colorQA);
				else
					sheet->SetColor(newWorksheet, row, 5 + i, color);
				sheet->SetStr(newWorksheet, ColIntToStr(5 + i) + (row + 1) + ":" + ColIntToStr(5 + i) + (row + 2), "");
			}

			sheet->SetStr(newWorksheet, row, 2, "Dev");
			sheet->SetStr(newWorksheet, row + 1, 2, "QA");

			sheet->SetStr(newWorksheet, "B" + (row + 2) + ":C" + (row + 2), "Total");
			sheet->SetFontBold(newWorksheet, "B" + (row + 2) + ":D" + (row + 2), true);
			sheet->SetStr(newWorksheet, row + 2, 4, "=SUM(" + "D" + row + ":D" + (row + 1) + ")");

			auto lastNotQACol = ColIntToStr(5 + (developers->Count - countQA - 1));
			auto firstQACol = ColIntToStr(5 + (developers->Count - countQA + 1));
			sheet->SetStr(newWorksheet, row, 3, "=IF(COUNT(E" + (row + 1) + ":" + lastNotQACol + (row + 1) + ")>0,AVERAGE(E" + (row + 1) + ":" + lastNotQACol + (row + 1) + "),0)");
			sheet->SetStr(newWorksheet, row + 1, 3, "=IF(COUNT(" + firstQACol + (row + 1) + ":" + lastColLetter + (row + 1) + ")>0,AVERAGE(" + firstQACol + (row + 1) + ":" + lastColLetter + (row + 1) + "),0)");

			auto storyRange = "A" + (row)+":A" + (row + 2);
			sheet->SetColor(newWorksheet, storyRange, color);

			sheet->SetBorder(newWorksheet, storyRange, true);
			sheet->SetBorder(newWorksheet, "B" + (row)+":D" + (row + 2), true);
			sheet->SetBorder(newWorksheet, "E" + (row)+":" + lastColLetter + (row + 2), true);
			sheet->SetBorder(newWorksheet, "E" + (row)+":" + lastColLetter + (row), true);
		}

		auto worksheet = sheet->GetWorksheetsByName("Focus factor")[0];
		int newCol = 2;
		for (; ; newCol++) {
			if (String::IsNullOrEmpty(sheet->GetStr(worksheet, 1, newCol))) {
				newCol--;
				sheet->InsertColLeft(worksheet, newCol);
				sheet->SetStr(worksheet, 1, newCol, newWorksheet->Name);
				sheet->SetHorAlign(worksheet, 1, newCol, XlHAlign::xlHAlignCenter);
				sheet->SetVerAlign(worksheet, 1, newCol, XlHAlign::xlHAlignCenter);
				sheet->SetColor(worksheet, 1, newCol, color);

				for (int row = 2; !String::IsNullOrEmpty(sheet->GetStr(worksheet, row, 1)); row++) {
					sheet->SetStr(worksheet, row, newCol, "-");
				}
				break;
			}
		}

		worksheet = sheet->GetWorksheetsByName("Sprint")[0];
		sheet->InsertRowUp(worksheet, 37);
		sheet->SetStr(worksheet, 37, 1, newWorksheet->Name);
		sheet->SetBorder(worksheet, 37, 1, true);
		sheet->SetStr(worksheet, 37, 2, "Focus factor");
		sheet->SetStr(worksheet, 37, 3, "Velocity");
		sheet->SetStr(worksheet, "D37:E37", "Absence");
		sheet->SetColor(worksheet, 37, 1, color);
		for (int i = 38; i < 38 + developers->Count; i++) {
			sheet->InsertRowUp(worksheet, i);
		}
		for (int i = 38; i < 38 + developers->Count; i++) {
			sheet->SetStr(worksheet, "D" + i + ":E" + i, "");
			sheet->SetStr(worksheet, i, 2, "=C" + i + "/$B$2");
			sheet->SetStr(worksheet, i, 3, i == 38 ? "0" : "9");

			sheet->SetBorder(worksheet, i, 1, false);
			sheet->SetBorder(worksheet, i, 2, false);
			sheet->SetBorder(worksheet, i, 3, false);
			sheet->SetBorder(worksheet, i, 4, false);
			sheet->SetBorder(worksheet, i, 5, false);
		}

		for (int i = 0; i < developers->Count; i++)
			sheet->SetStr(worksheet, 38 + i, 1, developers[i]);

		auto prelastRowNum = "" + (38 + developers->Count - 1);
		auto lastRowNum = "" + (38 + developers->Count);
		sheet->SetBorder(worksheet, "A" + lastRowNum + ":A" + lastRowNum, true);
		sheet->SetBorder(worksheet, "B" + lastRowNum + ":B" + lastRowNum, true);
		sheet->SetBorder(worksheet, "C" + lastRowNum + ":C" + lastRowNum, true);
		sheet->SetBorder(worksheet, "D" + lastRowNum + ":E" + lastRowNum, true);

		sheet->SetBorder(worksheet, "A38:A" + prelastRowNum, true);
		sheet->SetBorder(worksheet, "B38:B" + prelastRowNum, true);
		sheet->SetBorder(worksheet, "C38:C" + prelastRowNum, true);
		sheet->SetBorder(worksheet, "D38:E" + prelastRowNum, true);
		sheet->SetFontBold(worksheet, "A37:D37", true);
		sheet->SetHorAlign(worksheet, "A37:A37", XlHAlign::xlHAlignRight);
		sheet->SetHorAlign(worksheet, "B37:E37", XlHAlign::xlHAlignCenter);
		sheet->SetHorAlign(worksheet, "A38:C" + prelastRowNum, XlHAlign::xlHAlignRight);
		sheet->SetVerAlign(worksheet, "A37:E" + prelastRowNum, XlHAlign::xlHAlignCenter);

		auto totalStatRow = 1;
		while (!String::Equals(sheet->GetStr(worksheet, totalStatRow, 1), "Available human days:"))
			totalStatRow++;
		totalStatRow--;

		sheet->InsertColLeft(worksheet, 6);
		sheet->SetColor(worksheet, totalStatRow, 6, color);
		sheet->SetStr(worksheet, totalStatRow, 6, newWorksheet->Name);
		sheet->SetStr(worksheet, totalStatRow + 1, 6, "=SUM(C38:C" + prelastRowNum + ")");
		sheet->SetStr(worksheet, totalStatRow + 2, 6, "='Focus factor'!G56");
		sheet->SetStr(worksheet, totalStatRow + 3, 6, "=ROUND(F" + (totalStatRow + 1) + "*F" + (totalStatRow + 2) + ", 0)");
		sheet->SetStr(worksheet, totalStatRow + 4, 6, "='" + newWorksheet->Name + "'!D2");
		sheet->SetStr(worksheet, totalStatRow + 5, 6, "=F" + (totalStatRow + 4) + "/F" + (totalStatRow + 1));
		sheet->SetStr(worksheet, totalStatRow + 6, 6, "=F" + (totalStatRow + 3) + "-F" + (totalStatRow + 4));

		sheet->SetStr(newWorksheet, "G2:G2", "='Sprint'!F" + (totalStatRow + 5));
		sheet->SetStr(newWorksheet, "J2:J2", "='Sprint'!F" + (totalStatRow + 6));

		sheet->SetVisible(true);
	}
	};
}
