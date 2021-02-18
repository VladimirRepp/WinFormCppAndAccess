#include "MyForm.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;

[STAThread]
int main(array<String^>^ arg) {
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);

	ExampleCPP::MyForm form;
	Application::Run(% form);
}

System::Void ExampleCPP::MyForm::выходToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{	
	Application::Exit();
}

System::Void ExampleCPP::MyForm::обПрограммеToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
	MessageBox::Show("Данная программа показывает примеры работы на Windows Forms c++ и ACCESS.", "Внимание!");

	return System::Void();
}

System::Void ExampleCPP::MyForm::button_download_Click(System::Object^ sender, System::EventArgs^ e)
{
	//Подключаемся к БД
	String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";//Строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем запрос к БД
	dbConnection->Open();//открываем соеденение
	String^ query = "SELECT * FROM [table_name]";//запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);//команда
	OleDbDataReader^ dbReader = dbComand->ExecuteReader();//считываем данные

	//Проверяем данные
	if (dbReader->HasRows == false) {
		MessageBox::Show("Ошибка считывания данных!", "Ошибка!");
	}
	else {
		//Заполним данные в таблицу формы
		while (dbReader->Read()) {
			dataGridView1->Rows->Add(dbReader["id"], dbReader["Name"], dbReader["Cost"], dbReader["Quantity"]);
		}
	}

	//Закрываем соеденение
	dbReader->Close();
	dbConnection->Close();

	return System::Void();
}

System::Void ExampleCPP::MyForm::button_add_Click(System::Object^ sender, System::EventArgs^ e)
{
	//Выбор нужной строки для добавления
	if (dataGridView1->SelectedRows->Count != 1) {
		MessageBox::Show("Выберите одну строку для добавления!","Внимание!");
		return;
	}

	//Узанаем индекс выбранной строки
	int index = dataGridView1->SelectedRows[0]->Index;

	//Проверяем данные
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[3]->Value == nullptr) {
		MessageBox::Show("Не все данные введены!", "Внимание!");
		return;
	}

	//Считываем данные
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
	String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
	String^ cost = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
	String^ quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();

	//Подключаемся к БД
	String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";//Строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем запрос к БД
	dbConnection->Open();//открываем соеденение
	String^ query = "INSERT INTO [table_name] VALUES ("+ id +", '"+ name +"', "+ cost +", "+ quantity +")";//запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);//команда
	
	//Выполняем запрос
	if (dbComand->ExecuteNonQuery() != 1)
		MessageBox::Show("Ошибка выполнения запроса!","Ошибка!");
	else
		MessageBox::Show("Данные добавлены!", "Готово!");

	//Закрываем соеденение с БД
	dbConnection->Close();

	return System::Void();
}

System::Void ExampleCPP::MyForm::button_update_Click(System::Object^ sender, System::EventArgs^ e)
{
	//Выбор нужной строки для добавления
	if (dataGridView1->SelectedRows->Count != 1) {
		MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
		return;
	}

	//Узанаем индекс выбранной строки
	int index = dataGridView1->SelectedRows[0]->Index;

	//Проверяем данные
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
		dataGridView1->Rows[index]->Cells[3]->Value == nullptr) {
		MessageBox::Show("Не все данные введены!", "Внимание!");
		return;
	}

	//Считываем данные
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();
	String^ name = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
	String^ cost = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
	String^ quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();

	//Подключаемся к БД
	String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";//Строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем запрос к БД
	dbConnection->Open();//открываем соеденение
	String^ query = "UPDATE [table_name] SET Name='"+ name +"', Cost = "+ cost +", Quantity="+ quantity +" WHERE id = " + id;//запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);//команда

	//Выполняем запрос
	if (dbComand->ExecuteNonQuery() != 1)
		MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
	else
		MessageBox::Show("Данные изменены!", "Готово!");

	//Закрываем соеденение с БД
	dbConnection->Close();

	return System::Void();
}

System::Void ExampleCPP::MyForm::button_delete_Click(System::Object^ sender, System::EventArgs^ e)
{
	//Выбор нужной строки для добавления
	if (dataGridView1->SelectedRows->Count != 1) {
		MessageBox::Show("Выберите одну строку для добавления!", "Внимание!");
		return;
	}

	//Узанаем индекс выбранной строки
	int index = dataGridView1->SelectedRows[0]->Index;

	//Проверяем данные
	if (dataGridView1->Rows[index]->Cells[0]->Value == nullptr)
	{
		MessageBox::Show("Не все данные введены!", "Внимание!");
		return;
	}

	//Считываем данные
	String^ id = dataGridView1->Rows[index]->Cells[0]->Value->ToString();

	//Подключаемся к БД
	String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";//Строка подключения
	OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

	//Выполняем запрос к БД
	dbConnection->Open();//открываем соеденение
	String^ query = "DELETE FROM [table_name] WHERE id = " + id;//запрос
	OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);//команда

	//Выполняем запрос
	if (dbComand->ExecuteNonQuery() != 1)
		MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
	else {
		MessageBox::Show("Данные удалены!", "Готово!");
		dataGridView1->Rows->RemoveAt(index);//удалаяем строку из таблицы формы
	}

	//Закрываем соеденение с БД
	dbConnection->Close();

	return System::Void();
}
