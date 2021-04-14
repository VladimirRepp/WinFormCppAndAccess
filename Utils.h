#pragma once

// Required libraries
using namespace System;
using namespace System::Collections::Generic;
using namespace System::Windows::Forms;
using namespace System::Data::OleDb;

// Managed Class - as a structure with data
ref class Field {
public:
	int id;
	String^ name;
};

// Managed Class
ref class Utils
{
	// Fields
private: 
	List<Field^>^ fields; 

	// Methods
public:
	Utils();	// default constructor - creates an empty object
	~Utils();	// destructor - deletes
	void Download();
	void Insert();
	void Insert(Field^);
	void Update(Field^);
	void Update();
	void Delete(Field^);
	void Delete();
};
