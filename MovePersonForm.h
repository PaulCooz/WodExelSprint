#pragma once
#include "Sheet.h"

namespace WodExelSprint {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for MovePersonForm
	/// </summary>
	ref class MovePersonForm : public System::Windows::Forms::Form
	{
	private:
		Sheet^ sheet;

	public:
		MovePersonForm(Sheet^ s)
		{
			sheet = s;

			InitializeComponent();

			auto worksheets = sheet->GetWorksheetsByName("(.*team)|(DevOps)");
			comboBoxLeft->DropDownStyle = ComboBoxStyle::DropDownList;
			comboBoxRight->DropDownStyle = ComboBoxStyle::DropDownList;
			for (auto i = 0; i < worksheets->Count; i++)
			{
				comboBoxLeft->Items->Add(worksheets[i]->Name);
				comboBoxRight->Items->Add(worksheets[i]->Name);
			}

			dataGridViewLeft->RowCount = 1;
			dataGridViewLeft->Columns[0]->Name = "Developers";
			dataGridViewLeft->Columns[0]->AutoSizeMode = DataGridViewAutoSizeColumnMode::Fill;
			dataGridViewRight->RowCount = 1;
			dataGridViewRight->Columns[0]->Name = "Developers";
			dataGridViewRight->Columns[0]->AutoSizeMode = DataGridViewAutoSizeColumnMode::Fill;
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~MovePersonForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^ dataGridViewLeft;
	private: System::Windows::Forms::DataGridView^ dataGridViewRight;
	protected:

	protected:

	private: System::Windows::Forms::ComboBox^ comboBoxLeft;
	private: System::Windows::Forms::ComboBox^ comboBoxRight;
	private: System::Windows::Forms::Button^ buttonToRight;
	private: System::Windows::Forms::Button^ buttonToLeft;
	private: System::Windows::Forms::Button^ buttonOk;








	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataGridViewLeft = (gcnew System::Windows::Forms::DataGridView());
			this->dataGridViewRight = (gcnew System::Windows::Forms::DataGridView());
			this->comboBoxLeft = (gcnew System::Windows::Forms::ComboBox());
			this->comboBoxRight = (gcnew System::Windows::Forms::ComboBox());
			this->buttonToRight = (gcnew System::Windows::Forms::Button());
			this->buttonToLeft = (gcnew System::Windows::Forms::Button());
			this->buttonOk = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewLeft))->BeginInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewRight))->BeginInit();
			this->SuspendLayout();
			// 
			// dataGridViewLeft
			// 
			this->dataGridViewLeft->AllowUserToAddRows = false;
			this->dataGridViewLeft->AllowUserToDeleteRows = false;
			this->dataGridViewLeft->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridViewLeft->Location = System::Drawing::Point(12, 42);
			this->dataGridViewLeft->Name = L"dataGridViewLeft";
			this->dataGridViewLeft->ReadOnly = true;
			this->dataGridViewLeft->RowHeadersWidth = 51;
			this->dataGridViewLeft->RowTemplate->Height = 24;
			this->dataGridViewLeft->Size = System::Drawing::Size(290, 290);
			this->dataGridViewLeft->TabIndex = 0;
			// 
			// dataGridViewRight
			// 
			this->dataGridViewRight->AllowUserToAddRows = false;
			this->dataGridViewRight->AllowUserToDeleteRows = false;
			this->dataGridViewRight->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridViewRight->Location = System::Drawing::Point(343, 42);
			this->dataGridViewRight->Name = L"dataGridViewRight";
			this->dataGridViewRight->ReadOnly = true;
			this->dataGridViewRight->RowHeadersWidth = 51;
			this->dataGridViewRight->RowTemplate->Height = 24;
			this->dataGridViewRight->Size = System::Drawing::Size(290, 290);
			this->dataGridViewRight->TabIndex = 1;
			// 
			// comboBoxLeft
			// 
			this->comboBoxLeft->FormattingEnabled = true;
			this->comboBoxLeft->Location = System::Drawing::Point(12, 12);
			this->comboBoxLeft->Name = L"comboBoxLeft";
			this->comboBoxLeft->Size = System::Drawing::Size(290, 24);
			this->comboBoxLeft->TabIndex = 2;
			this->comboBoxLeft->SelectedIndexChanged += gcnew System::EventHandler(this, &MovePersonForm::comboBoxLeft_SelectedIndexChanged);
			// 
			// comboBoxRight
			// 
			this->comboBoxRight->FormattingEnabled = true;
			this->comboBoxRight->Location = System::Drawing::Point(343, 12);
			this->comboBoxRight->Name = L"comboBoxRight";
			this->comboBoxRight->Size = System::Drawing::Size(290, 24);
			this->comboBoxRight->TabIndex = 3;
			this->comboBoxRight->SelectedIndexChanged += gcnew System::EventHandler(this, &MovePersonForm::comboBoxRight_SelectedIndexChanged);
			// 
			// buttonToRight
			// 
			this->buttonToRight->Location = System::Drawing::Point(308, 165);
			this->buttonToRight->Name = L"buttonToRight";
			this->buttonToRight->Size = System::Drawing::Size(29, 23);
			this->buttonToRight->TabIndex = 4;
			this->buttonToRight->Text = L">";
			this->buttonToRight->UseVisualStyleBackColor = true;
			this->buttonToRight->Click += gcnew System::EventHandler(this, &MovePersonForm::buttonToRight_Click);
			// 
			// buttonToLeft
			// 
			this->buttonToLeft->Location = System::Drawing::Point(308, 194);
			this->buttonToLeft->Name = L"buttonToLeft";
			this->buttonToLeft->Size = System::Drawing::Size(29, 23);
			this->buttonToLeft->TabIndex = 5;
			this->buttonToLeft->Text = L"<";
			this->buttonToLeft->UseVisualStyleBackColor = true;
			this->buttonToLeft->Click += gcnew System::EventHandler(this, &MovePersonForm::buttonToLeft_Click);
			// 
			// buttonOk
			// 
			this->buttonOk->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 14));
			this->buttonOk->Location = System::Drawing::Point(12, 354);
			this->buttonOk->Name = L"buttonOk";
			this->buttonOk->Size = System::Drawing::Size(621, 48);
			this->buttonOk->TabIndex = 6;
			this->buttonOk->Text = L"apply";
			this->buttonOk->UseVisualStyleBackColor = true;
			this->buttonOk->Click += gcnew System::EventHandler(this, &MovePersonForm::buttonOk_Click);
			// 
			// MovePersonForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(645, 414);
			this->Controls->Add(this->buttonOk);
			this->Controls->Add(this->buttonToLeft);
			this->Controls->Add(this->buttonToRight);
			this->Controls->Add(this->comboBoxRight);
			this->Controls->Add(this->comboBoxLeft);
			this->Controls->Add(this->dataGridViewRight);
			this->Controls->Add(this->dataGridViewLeft);
			this->Name = L"MovePersonForm";
			this->Text = L"Move Person";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewLeft))->EndInit();
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridViewRight))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void comboBoxLeft_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
		dataGridViewLeft->Rows->Clear();
		auto worksheet = sheet->GetWorksheetsByName(comboBoxLeft->Text)[0];
		for (auto i = 5; ; i++) {
			auto name = sheet->GetStr(worksheet, 4, i);
			if (String::IsNullOrEmpty(name))
				break;
			dataGridViewLeft->Rows->Add(name);
		}
	}
	private: System::Void comboBoxRight_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
		dataGridViewRight->Rows->Clear();
		auto worksheet = sheet->GetWorksheetsByName(comboBoxRight->Text)[0];
		for (auto i = 5; ; i++) {
			auto name = sheet->GetStr(worksheet, 4, i);
			if (String::IsNullOrEmpty(name))
				break;
			dataGridViewRight->Rows->Add(name);
		}
	}
	private: System::Void buttonToRight_Click(System::Object^ sender, System::EventArgs^ e) {
		auto index = dataGridViewLeft->CurrentCell->RowIndex;
		if (0 <= index && index < dataGridViewLeft->Rows->Count)
		{
			auto name = (String^)(dataGridViewLeft->CurrentCell->Value);
			dataGridViewLeft->Rows->RemoveAt(index);

			if (name->StartsWith("[QA]")) {
				dataGridViewRight->Rows->Add(name);
			}
			else {
				auto lastNotQA = -1;
				for (auto i = 0; i < dataGridViewRight->Rows->Count; i++) {
					auto name = (String^)(dataGridViewRight->Rows[i]->Cells[0]->Value);
					if (!name->StartsWith("[QA]")) {
						lastNotQA = i;
					}
				}
				lastNotQA++;
				dataGridViewRight->Rows->Insert(lastNotQA);
				dataGridViewRight->Rows[lastNotQA]->Cells[0]->Value = name;
			}
		}
	}
	private: System::Void buttonToLeft_Click(System::Object^ sender, System::EventArgs^ e) {
		auto index = dataGridViewRight->CurrentCell->RowIndex;
		if (0 <= index && index < dataGridViewRight->Rows->Count)
		{
			auto name = (String^)(dataGridViewRight->CurrentCell->Value);
			dataGridViewRight->Rows->RemoveAt(index);

			if (name->StartsWith("[QA]")) {
				dataGridViewLeft->Rows->Add(name);
			}
			else {
				auto lastNotQA = -1;
				for (auto i = 0; i < dataGridViewLeft->Rows->Count; i++) {
					auto name = (String^)(dataGridViewLeft->Rows[i]->Cells[0]->Value);
					if (!name->StartsWith("[QA]")) {
						lastNotQA = i;
					}
				}
				lastNotQA++;
				dataGridViewLeft->Rows->Insert(lastNotQA);
				dataGridViewLeft->Rows[lastNotQA]->Cells[0]->Value = name;
			}
		}
	}
	private: System::Void buttonOk_Click(System::Object^ sender, System::EventArgs^ e) {
		Close();
	}
	public: String^ GetLeftTeam() {
		return comboBoxLeft->Text;
	}
	public: List<String^ >^ GetLeftTeammates() {
		auto res = gcnew List<String^>();
		for (auto i = 0; i < dataGridViewLeft->Rows->Count; i++) {
			res->Add((String^)dataGridViewLeft->Rows[i]->Cells[0]->Value);
		}
		return res;
	}
	public: String^ GetRightTeam() {
		return comboBoxRight->Text;
	}
	public: List<String^ >^ GetRightTeammates() {
		auto res = gcnew List<String^>();
		for (auto i = 0; i < dataGridViewRight->Rows->Count; i++) {
			res->Add((String^)dataGridViewRight->Rows[i]->Cells[0]->Value);
		}
		return res;
	}
	};
}
