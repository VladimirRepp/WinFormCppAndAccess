#include "Utils.h"

Utils::Utils()
{
	fields = gcnew List<Field^>();
}

Utils::~Utils()
{
	fields->Clear();
}

void Utils::Download()
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb"; // connection string
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open();									// opening the connection
		String^ query = "SELECT * FROM [table_name]";
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
		OleDbDataReader^ dbReader = dbComand->ExecuteReader();	// reading the data

		// Checking the data
		if (dbReader->HasRows == false) {
			MessageBox::Show("Data reading error!", "Error!");
		}
		else {
			fields->Clear();

			// We do the necessary actions with the data
			while (dbReader->Read()) {
				// For example:  
				Field^ buf = gcnew Field();
				buf->id = Convert::ToInt32(dbReader["id"]);
				buf->name = Convert::ToString(dbReader["Name"]);

				fields->Add(buf);
			}
		}

		// Closing the connection
		dbReader->Close();
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}

void Utils::Insert()
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb"; // connection string
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open(); // opening the connection

		// Insert all data
		bool err = false;
		for (int i = 0; i < fields->Count; i++) {
			String^ query = "INSERT INTO [table_name] VALUES (" + fields[i]->id + ", '" + fields[i]->name + "')";
			OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

			// Executing the request
			if (dbComand->ExecuteNonQuery() != 1)
				err = true;
		}

		if (err) {
			MessageBox::Show("Error inserting!", "Error!");
		}
		else{
			MessageBox::Show("All data inserting!!", "Ошибка!");
		}

		// Closing the connection
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}

void Utils::Insert(Field^ f)
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb"; // connection string
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open(); // opening the connection

		String^ query = "INSERT INTO [table_name] VALUES (" + f->id + ", '" + f->name + "')";
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

		// Executing the request
		if (dbComand->ExecuteNonQuery() != 1) {
			MessageBox::Show("Error inserting!", "Error!");
		}
		else {
			MessageBox::Show("All data inserting!!", "Ошибка!");
		}

		// Closing the connection
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}

void Utils::Update(Field^ f)
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open();
		String^ query = "UPDATE [table_name] SET Name='" + f->name + " WHERE id = " + f->id;
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

		// Executing the request
		if (dbComand->ExecuteNonQuery() != 1)
			MessageBox::Show("Updating error!", "Error!");
		else
			MessageBox::Show("Data update!", "Done!");

		// Closing the connection to the database
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}

void Utils::Update()
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open();

		// Update all data
		bool err = false;
		for each (Field^ f in fields)
		{
			String^ query = "UPDATE [table_name] SET Name='" + f->name + " WHERE id = " + f->id;
			OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

			// Executing the request
			if (dbComand->ExecuteNonQuery() != 1)
				err = true;
		}
		
		if(err)
			MessageBox::Show("Updating error!", "Error!");
		else
			MessageBox::Show("Data update!", "Done!");

		// Closing the connection to the database
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}

void Utils::Delete(Field^ f)
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open();
		String^ query = "DELETE FROM [table_name] WHERE id = " + f->id;
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

		// Executing the request
		if (dbComand->ExecuteNonQuery() != 1)
			MessageBox::Show("Deletion error!", "Error!");
		else {
			MessageBox::Show("Data deleted!", "Done!");
		}

		// Closing the connection
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}

void Utils::Delete()
{
	try {
		// Connecting to the database
		String^ connectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb";
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		// Executing a query to the database
		dbConnection->Open();
		String^ query = "DELETE FROM [table_name]"; // clear table
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);

		// Executing the request
		if (dbComand->ExecuteNonQuery() != 1)
			MessageBox::Show("Deletion error!", "Error!");
		else {
			MessageBox::Show("Data deleted!", "Done!");
		}

		// Closing the connection
		dbConnection->Close();
	}
	catch (OleDbException^ e) {
		MessageBox::Show(e->ToString(), "Error!");
	}
}
