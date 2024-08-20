#pragma once

namespace WodExelSprint {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Collections::Generic;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Summary for AddTeamForm
	/// </summary>
	public ref class AddTeamForm : public System::Windows::Forms::Form
	{
	public:
		AddTeamForm(void)
		{
			InitializeComponent();

			dataGridView1->RowCount = 1;
			dataGridView1->Columns[0]->Name = "Developers";
			dataGridView1->Columns[0]->AutoSizeMode = DataGridViewAutoSizeColumnMode::Fill;
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~AddTeamForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::TextBox^ textBox1;
	private: System::Windows::Forms::Label^ label_team;
	private: System::Windows::Forms::Label^ label_color;
	private: System::Windows::Forms::Label^ label_color_value;
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::Button^ button2;
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::Button^ append;
	private: System::Windows::Forms::Button^ pop_back;







	protected:


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
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle1 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			System::Windows::Forms::DataGridViewCellStyle^ dataGridViewCellStyle2 = (gcnew System::Windows::Forms::DataGridViewCellStyle());
			this->textBox1 = (gcnew System::Windows::Forms::TextBox());
			this->label_team = (gcnew System::Windows::Forms::Label());
			this->label_color = (gcnew System::Windows::Forms::Label());
			this->label_color_value = (gcnew System::Windows::Forms::Label());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->button2 = (gcnew System::Windows::Forms::Button());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->append = (gcnew System::Windows::Forms::Button());
			this->pop_back = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// textBox1
			// 
			this->textBox1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Right));
			this->textBox1->Location = System::Drawing::Point(195, 15);
			this->textBox1->Name = L"textBox1";
			this->textBox1->Size = System::Drawing::Size(332, 22);
			this->textBox1->TabIndex = 0;
			this->textBox1->TextChanged += gcnew System::EventHandler(this, &AddTeamForm::textBox1_TextChanged);
			// 
			// label_team
			// 
			this->label_team->AutoSize = true;
			this->label_team->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label_team->Location = System::Drawing::Point(13, 12);
			this->label_team->Name = L"label_team";
			this->label_team->Size = System::Drawing::Size(115, 25);
			this->label_team->TabIndex = 1;
			this->label_team->Text = L"team name:";
			// 
			// label_color
			// 
			this->label_color->AutoSize = true;
			this->label_color->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label_color->Location = System::Drawing::Point(12, 60);
			this->label_color->Name = L"label_color";
			this->label_color->Size = System::Drawing::Size(108, 25);
			this->label_color->TabIndex = 2;
			this->label_color->Text = L"team color:";
			// 
			// label_color_value
			// 
			this->label_color_value->AutoSize = true;
			this->label_color_value->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->label_color_value->Location = System::Drawing::Point(190, 60);
			this->label_color_value->Name = L"label_color_value";
			this->label_color_value->Size = System::Drawing::Size(95, 25);
			this->label_color_value->TabIndex = 3;
			this->label_color_value->Text = L"#FFFFFF";
			// 
			// button1
			// 
			this->button1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Right));
			this->button1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button1->Location = System::Drawing::Point(371, 52);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(156, 33);
			this->button1->TabIndex = 4;
			this->button1->Text = L"choose";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &AddTeamForm::button1_Click);
			// 
			// button2
			// 
			this->button2->Anchor = System::Windows::Forms::AnchorStyles::Bottom;
			this->button2->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->button2->Location = System::Drawing::Point(165, 376);
			this->button2->Name = L"button2";
			this->button2->Size = System::Drawing::Size(181, 39);
			this->button2->TabIndex = 5;
			this->button2->Text = L"apply";
			this->button2->UseVisualStyleBackColor = true;
			this->button2->Click += gcnew System::EventHandler(this, &AddTeamForm::button2_Click);
			// 
			// dataGridView1
			// 
			this->dataGridView1->AllowUserToAddRows = false;
			this->dataGridView1->AllowUserToDeleteRows = false;
			this->dataGridView1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->dataGridView1->BackgroundColor = System::Drawing::Color::White;
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			dataGridViewCellStyle1->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleLeft;
			dataGridViewCellStyle1->BackColor = System::Drawing::SystemColors::Window;
			dataGridViewCellStyle1->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 10.2F, System::Drawing::FontStyle::Regular,
				System::Drawing::GraphicsUnit::Point, static_cast<System::Byte>(0)));
			dataGridViewCellStyle1->ForeColor = System::Drawing::SystemColors::ControlText;
			dataGridViewCellStyle1->SelectionBackColor = System::Drawing::SystemColors::Highlight;
			dataGridViewCellStyle1->SelectionForeColor = System::Drawing::SystemColors::HighlightText;
			dataGridViewCellStyle1->WrapMode = System::Windows::Forms::DataGridViewTriState::False;
			this->dataGridView1->DefaultCellStyle = dataGridViewCellStyle1;
			this->dataGridView1->Location = System::Drawing::Point(18, 143);
			this->dataGridView1->Name = L"dataGridView1";
			dataGridViewCellStyle2->Alignment = System::Windows::Forms::DataGridViewContentAlignment::MiddleLeft;
			dataGridViewCellStyle2->BackColor = System::Drawing::SystemColors::Control;
			dataGridViewCellStyle2->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			dataGridViewCellStyle2->ForeColor = System::Drawing::SystemColors::WindowText;
			dataGridViewCellStyle2->SelectionBackColor = System::Drawing::SystemColors::Highlight;
			dataGridViewCellStyle2->SelectionForeColor = System::Drawing::SystemColors::HighlightText;
			dataGridViewCellStyle2->WrapMode = System::Windows::Forms::DataGridViewTriState::True;
			this->dataGridView1->RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
			this->dataGridView1->RowHeadersWidth = 51;
			this->dataGridView1->RowTemplate->Height = 24;
			this->dataGridView1->Size = System::Drawing::Size(509, 191);
			this->dataGridView1->TabIndex = 6;
			// 
			// append
			// 
			this->append->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 10.2F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->append->Location = System::Drawing::Point(437, 340);
			this->append->Name = L"append";
			this->append->Size = System::Drawing::Size(43, 30);
			this->append->TabIndex = 7;
			this->append->Text = L"+";
			this->append->UseVisualStyleBackColor = true;
			this->append->Click += gcnew System::EventHandler(this, &AddTeamForm::append_Click);
			// 
			// pop_back
			// 
			this->pop_back->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 10.2F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(0)));
			this->pop_back->Location = System::Drawing::Point(486, 340);
			this->pop_back->Name = L"pop_back";
			this->pop_back->Size = System::Drawing::Size(41, 30);
			this->pop_back->TabIndex = 8;
			this->pop_back->Text = L"-";
			this->pop_back->UseVisualStyleBackColor = true;
			this->pop_back->Click += gcnew System::EventHandler(this, &AddTeamForm::pop_back_Click);
			// 
			// AddTeamForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(539, 427);
			this->Controls->Add(this->pop_back);
			this->Controls->Add(this->append);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->button2);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->label_color_value);
			this->Controls->Add(this->label_color);
			this->Controls->Add(this->label_team);
			this->Controls->Add(this->textBox1);
			this->Name = L"AddTeamForm";
			this->Text = L"AddTeamForm";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	public:
		System::Drawing::Color color = System::Drawing::Color::White;
		String^ teamName = "-";

	private: System::Void button2_Click(System::Object^ sender, System::EventArgs^ e) {
		this->Close();
	}
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		auto colorDialog = gcnew ColorDialog();
		if (colorDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
			color = colorDialog->Color;
		label_color_value->Text = ColorTranslator::ToHtml(color);
	}
	private: System::Void textBox1_TextChanged(System::Object^ sender, System::EventArgs^ e) {
		teamName = textBox1->Text;
	}
	public: List<String^>^ GetDevelopers() {
		auto res = gcnew List<String^>();
		for (int i = 0; i < dataGridView1->RowCount; i++) {
			if (dataGridView1->Rows[i]->Cells->Count == 0)
				continue;
			auto val = (String^)(dataGridView1->Rows[i]->Cells[0]->Value);
			if (String::IsNullOrEmpty(val))
				continue;

			res->Add(val);
		}
		return res;
	}
	private: System::Void append_Click(System::Object^ sender, System::EventArgs^ e) {
		dataGridView1->Rows->Add();
	}
	private: System::Void pop_back_Click(System::Object^ sender, System::EventArgs^ e) {
		auto index = dataGridView1->Rows->Count - 2;
		auto row = dataGridView1->CurrentCell->RowIndex;
		if (0 <= row && row <= index)
			index = row;

		if (index >= 0)
			dataGridView1->Rows->RemoveAt(index);
	}
	};
}
