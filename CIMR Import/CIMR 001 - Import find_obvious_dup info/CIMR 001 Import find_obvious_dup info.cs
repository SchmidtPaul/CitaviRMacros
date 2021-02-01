using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;			
using System.Data.OleDb;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using SwissAcademic.Citavi;
using SwissAcademic.Citavi.DataExchange;
using SwissAcademic.Citavi.Metadata;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;

/// Make sure you have the Microsoft.ACE.OLEDB.16.0 provider installed, which is part of the Microsoft Access Database Engine 2016 Redistributable:
/// https://www.microsoft.com/en-us/download/details.aspx?id=54920
/// Otherweise you will receive an error saying "Microsoft.ACE.OLEDB.16.0 provider is not registered on the local machine".

public static class CitaviMacro{

	public static void Main(){


	// user settings for import ------------------------------------------------

		// define names of excel columns to be read from
		string[] excelColNames = new string[100]; // dont change
		excelColNames[0] = "ID"; // excel column name to be merged by (=Citavi-ID)
		excelColNames[1] = "clean_title";    // excel column to be read from #1
		excelColNames[2] = "clean_title_id"; // excel column to be read from #2      
		excelColNames[3] = "obv_dup_id";     // excel column to be read from #3
		// excelColNames[4] = "has_obv_dup";    // excel column to be read from #4

		// define names of citavi fields to be written into
		ReferencePropertyId[] citaviField = new ReferencePropertyId[4];
		citaviField[0] = ReferencePropertyId.TranslatedTitle; // dummy filler, never used, no need to change
		citaviField[1] = ReferencePropertyId.CustomField1;    // citavi field to be written into #1
		citaviField[2] = ReferencePropertyId.CustomField2;    // citavi field to be written into #2
		citaviField[3] = ReferencePropertyId.CustomField3;    // citavi field to be written into #3
		// citaviField[4] = ReferencePropertyId.CustomField4;    // citavi field to be written into #4


		excelColNames = excelColNames.Where(x => !string.IsNullOrEmpty()).ToArray(); // dont change


	// read from excel columns (= excel import) --------------------------------

		Project project = Program.ActiveProjectShell.Project;	
		DataTable dataTable = new DataTable();
		string worksheetNameToImport = ""; // leave empty
		string fileName = ""; // leave empty
		int n_IDsSuccessfullReadWrite = 0; // will count

		// Window: get path of excel file
		using (OpenFileDialog dialog = new OpenFileDialog())
		{
			dialog.Filter = SwissAcademic.Resources.FileDialogFilters.Excel;
			dialog.InitialDirectory = @"C:\Users\<your name>\Desktop";
			dialog.Title = "Choose EXCEL file to import with data on first sheet";
			if (dialog.ShowDialog() == DialogResult.OK)
			{
				fileName = dialog.FileName;
			}
			else
			{
				return;
			}
		}

		DebugMacro.WriteLine(string.Format("Trying to read first worksheet from '" + fileName + "'"));

		// get name of excel sheet in excel file
		string sheetName = GetExistingSheetName(worksheetNameToImport, fileName);
		if (string.IsNullOrEmpty(sheetName)) return;
		DebugMacro.WriteLine(string.Format("   Name of imported worksheet is '" + sheetName + "'"));

		// read data on excel sheet in excel file
		dataTable = Sheet2DataTable(fileName, sheetName, -1); 
		if (dataTable == null) 
		{
			DebugMacro.WriteLine("   Error: No data could be read.");
			return;
		}
		else
		{
			DebugMacro.WriteLine("   Data was successfully read.");
		}


	// check user settings and excel file --------------------------------------

		DebugMacro.WriteLine("Checking if all excel columns defined in 'user settings for import' can be found in the read data:");

		DataColumn[] excelColData = new DataColumn[citaviField.Length];
		for (int i = 0; i < excelColData.Length; i++){excelColData[i] = null;}

		// Check if all columns defined in "user settings for import" above can be found in read excel columns
		foreach (DataColumn col in dataTable.Columns)
		{
			
			for (int i = 0; i < excelColData.Length; i++)
			{
				if (excelColData[i] == null && col.ToString() == excelColNames[i])
				{
					excelColData[i] = col;
					DebugMacro.WriteLine(string.Format("   Found '" + col + "'"));
				}
			}
		}

		// Error if not all columns defined in "user settings for import" above could be found in read excel columns
		for (int i = 0; i < excelColData.Length; i++)
		{
			if (excelColData[i] == null)
			{
				MessageBox.Show("   Error: Could not find required column " + excelColNames[i] + ".", 
					"Citavi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
		}

		DebugMacro.WriteLine("   All excel columns defined in 'user settings for import' found.");


	// write into citavi fields ------------------------------------------------

		DebugMacro.WriteLine(string.Format("Trying to write data into respective citavi fields."));
		string MergeColName = excelColNames[0];

		for (int i = 1; i < dataTable.Rows.Count; i++)
		{

			string MergeColEntry_i = dataTable.Rows[i][excelColData[0]].ToString();

			if (string.IsNullOrEmpty(MergeColEntry_i))
			{
				// writing impossible if ID is empty for a reference
				DebugMacro.WriteLine("   Not writing data for reference in row " + i + " of excel file because '" + MergeColName + "' is empty."); 
			}
			else
			{
				// determine which parentReference in citavi has the same ID
				Reference parentReference;
				parentReference = GetReferenceWithCitaviID(MergeColEntry_i);

				if (parentReference == null)
				{
					// if the 'parentReference' is not present in the Citavi project
					DebugMacro.WriteLine("   Not writing data for reference in row " + i + " of excel file because no reference with the same '" + MergeColName + "' exists in this Citavi project (" + MergeColEntry_i + ")");	
				}
				
				else
				{
					// if the 'parentReference' is present in the Citavi project
					DebugMacro.WriteLine("   Writing data for reference in row " + i + " of excel file into existing reference with the same '" + MergeColName + "' (" + MergeColEntry_i + ")");

					for (int j = 1; j < citaviField.Length; j++)
					{
						if (!string.IsNullOrEmpty(dataTable.Rows[i][excelColData[j]].ToString())) // only write if there is something to write.
						{
							// (over)write read data into all fields defined in "user settings for import" for this 'parentReference'
							parentReference.SetValue(citaviField[j], dataTable.Rows[i][excelColData[j]].ToString());
						}
					}						

					n_IDsSuccessfullReadWrite++;
				}
			}
		}


	// Summary -----------------------------------------------------------------

		int n_referencesInCitavi = project.References.Count;
		int n_referencesInExcel  = dataTable.Rows.Count - 1; // -1 for column header

		DebugMacro.WriteLine("===============");
		DebugMacro.WriteLine("     Finished Macro");
		DebugMacro.WriteLine("===============");
		DebugMacro.WriteLine("   " + n_referencesInCitavi + " references in citavi project");
		DebugMacro.WriteLine("   " + n_referencesInExcel + " references in read excel data");
		DebugMacro.WriteLine("   " + n_IDsSuccessfullReadWrite + " references were successfully read from excel and written into citavi:");

		for (int j = 1; j < citaviField.Length; j++)
		{
		DebugMacro.WriteLine("      " + excelColData[j] + " was written into " + citaviField[j]);
		}


	}


	// -------------------------------------------------------------------------
	// -------------------------------------------------------------------------
	// End of main part; Functions follow --------------------------------------
	// -------------------------------------------------------------------------
	// -------------------------------------------------------------------------


	// Find a reference via its Citavi-ID --------------------------------------
	private static Reference GetReferenceWithCitaviID(string CitaviID)
	{
		Project project = Program.ActiveProjectShell.Project;	

		foreach (Reference reference in project.References)
		{
			if (reference.Id.ToString() == CitaviID) return reference;				
		}
		return null;
	}

	// dealing with excel import. all copied from CIM007 -----------------------
	// https://github.com/Citavi/Macros/blob/master/CIM%20Import/CIM007%20Import%20arbitrary%20data%20from%20Microsoft%20Excel%20into%20custom%20fields%20of%20existing%20references%20by%20short%20title/CIM007_Import_Excel_Data.cs

		// get path of excel file
		private static string GetConnectionString(string fileName)
		{
			string connectionString = string.Empty;
			connectionString = "Provider=Microsoft.ACE.OLEDB.16.0;" +
			"Data Source={0};Extended Properties=" + Convert.ToChar(34).ToString() + "Excel 12.0;HDR=YES;" + Convert.ToChar(34).ToString();
			return string.Format(connectionString, fileName);
		}

		// get name of excel sheet in excel file
		private static string GetExistingSheetName(string requestedSheetName, string fileName) 
		{
			List<string> sheetNames = ExcelFetcher.GetWorksheets(fileName);
			if (sheetNames.Count == 0) return "";
			
			foreach (string sheetName in sheetNames)
			{
				if (sheetName == requestedSheetName) return sheetName;
			}
			return sheetNames[0];
		}

		// read data on excel sheet in excel file
		private static DataTable Sheet2DataTable(string fileName, string sheetName, int maxRowCount)
		{
			DataTable dataTable = new DataTable();
			OleDbDataReader dataReader = null;
			DataRow row = null;

			string selectString = @"SELECT * FROM ["
									+ "{0}"
									+ "]";

			try
			{
				using (OleDbConnection connection = new OleDbConnection(GetConnectionString(fileName)))
				{
					connection.Open();
					object[] o = new Object[] { null, null, null, "TABLE" };
					using (DataTableReader dataReader2 = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, o).CreateDataReader())
					{
						while (dataReader2.Read())
						{
							string sheetName2 = dataReader2["TABLE_NAME"].ToString();
							if (sheetName2.EndsWith("$") ||
								sheetName2.EndsWith("$_") ||
								sheetName2.EndsWith("$'") ||
								sheetName2.EndsWith("$'_"))
							{
								sheetName2 = sheetName2.Remove(sheetName2.IndexOf("$"));
							}
							if (sheetName2.StartsWith("'"))
							{
								sheetName2 = sheetName2.Remove(0, 1);
							}
							if (sheetName == sheetName2)
							{
								sheetName = dataReader2["TABLE_NAME"].ToString();
								break;
							}
						}
					}
				}
				
				//// DebugMacro.WriteLine(string.Format("   Trying to populate Datatable from Worksheet '{0}'", sheetName));

				using (OleDbConnection connection = new OleDbConnection(GetConnectionString(fileName)))
				{
					connection.Open();
					using (OleDbCommand command = new OleDbCommand(string.Format(selectString, sheetName), connection))
					{
						dataReader = command.ExecuteReader();

						while (dataReader.Read())
						{
							row = dataTable.NewRow();
							for (int i = 0; i < dataReader.FieldCount; i++)
							{
								if (dataTable.Columns.Count == i)
								{
									string name = dataReader.GetName(i);
									if (dataTable.Columns.Contains(name))
									{
										dataTable.Columns.Add("Column" + i.ToString());
									}
									else
									{
										dataTable.Columns.Add(name);
									}
								}
								if (dataReader[i] != DBNull.Value)
								{
									row[i] = dataReader[i].ToString();
								}
								else
								{
									row[i] = string.Empty;
								}
							}
							dataTable.Rows.Add(row);
							maxRowCount--;
							if (maxRowCount == 0)
								break;
						}
					}
				}

				row = dataTable.NewRow();
				foreach (DataColumn column in dataTable.Columns)
				{
					row[column] = column.ColumnName;
				}
				dataTable.Rows.InsertAt(row, 0);
				
				#region Clean

				
				bool isEmpty = true;
				for (int i = 0; i < dataTable.Columns.Count; i++)
				{
					isEmpty = true;
					for (int i1 = 0; i1 < dataTable.Rows.Count; i1++)
					{
						if (!string.IsNullOrEmpty(dataTable.Rows[i1][i].ToString()))
						{
							isEmpty = false;
							i1 = dataTable.Rows.Count;
						}
					}
						if (isEmpty)
					{
						dataTable.Columns.RemoveAt(i);
						i--;
					}
				}
				for (int i = 0; i < dataTable.Rows.Count; i++)
				{
					isEmpty = true;
					for (int i1 = 0; i1 < dataTable.Columns.Count; i1++)
					{
						if (!string.IsNullOrEmpty(dataTable.Rows[i][i1].ToString()))
						{
							isEmpty = false;
							i1 = dataTable.Columns.Count;
						}
					}
						if (isEmpty)
					{
						dataTable.Rows.RemoveAt(i);
						i--;
					}
				}
					#endregion
				}
			catch (Exception ignored)
			{
				MessageBox.Show(ignored.Message);
			}
			finally
			{
				if (dataReader != null && !dataReader.IsClosed)
					dataReader.Close();
			}
			return dataTable;
		}

	// -------------------------------------------------------------------------
	
}