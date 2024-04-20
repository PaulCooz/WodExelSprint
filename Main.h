#pragma once

#include "Checker.cpp"

namespace WodExelSprint {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::IO;
	using namespace System::Diagnostics;
	using namespace Microsoft::Office::Interop;

	/// <summary>
	/// Summary for Main
	/// </summary>
	public ref class Main : public System::Windows::Forms::Form
	{
	private: System::Windows::Forms::Button^ OpenXlsxButton;
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
			this->OpenXlsxButton = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// OpenXlsxButton
			// 
			this->OpenXlsxButton->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->OpenXlsxButton->Location = System::Drawing::Point(12, 12);
			this->OpenXlsxButton->Name = L"OpenXlsxButton";
			this->OpenXlsxButton->Size = System::Drawing::Size(209, 68);
			this->OpenXlsxButton->TabIndex = 0;
			this->OpenXlsxButton->Text = L"open xlsx";
			this->OpenXlsxButton->UseVisualStyleBackColor = true;
			this->OpenXlsxButton->Click += gcnew System::EventHandler(this, &Main::OpenXlsxFile);
			// 
			// Main
			// 
			this->ClientSize = System::Drawing::Size(233, 92);
			this->Controls->Add(this->OpenXlsxButton);
			this->Name = L"Main";
			this->ResumeLayout(false);

		}

	private: System::Void OpenXlsxFile(System::Object^ sender, System::EventArgs^ e) {
		OpenFileDialog^ openFileDialog = gcnew OpenFileDialog;

		openFileDialog->InitialDirectory = ".";
		openFileDialog->Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
		openFileDialog->FilterIndex = 1;
		openFileDialog->RestoreDirectory = true;

		if (openFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			Excel::Application^ xlApp = gcnew Excel::ApplicationClass();

			Excel::Workbook^ workbook = xlApp->Workbooks->Open(
				openFileDialog->FileName,
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

			CheckCells((Excel::Worksheet^)(workbook->ActiveSheet));

			xlApp->Visible = true;
		}
	}
	};
}
