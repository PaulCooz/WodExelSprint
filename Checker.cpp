using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Windows::Forms;
using namespace System::Data;
using namespace System::Drawing;
using namespace System::IO;
using namespace System::Diagnostics;
using namespace Microsoft::Office::Interop;

Void CheckCells(Excel::Worksheet^ excelWorksheet)
{
	auto range = excelWorksheet->Range["D5", "D5"];
	auto text = range->Text;
	auto cellValue = text->ToString();
	if (String::IsNullOrEmpty(cellValue)) {
		text = "error"; // TODO set error
	}
}
