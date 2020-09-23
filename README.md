<div align="center">

## Database Connection Class


</div>

### Description

This code serves as a starting point for database connections in C#. Adding greater functionality should not be too difficult.

This example contains functionality for data navigation, database connection, recordset connection, Bookmarking, AbsolutePosition and RecordCount counters, and of course: Returning data from the table.
 
### More Info
 
In order to use this code you will need to set a reference to the Microsoft ActiveX Data Objects library 2.7.

A quick example of how to use this code:

//Declare an object of type DBConnect class:

private DBConnect DBCN = new DBConnect();

//Because the connection is static (so you

//can have one connection but several recordsets)

//You connect to the database using the class name

//and not the object you declared:

DBConnect.Connect("\\Database\\ExampleDB.mdb");

//To open a recordset:

DBCN.RSConnect("tblTableName");

//To navigate records:

DBCN.Navigate("first");

DBCN.Navigate("next");

DBCN.Navigate("previous");

DBCN.Navigate("Last");

//To set a bookmark:

DBCN.Bookmark = DBCN.AbsPosition;

//To go to the bookmarked record:

DBCN.GotoBookmark();

//To return data from a field on

//the current record (0 is the number

//for the field):

DBCN.Value(0)

If you do not have the .NET framework installed on your machine, this code has no hope of working.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mandrake Dragonbait](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mandrake-dragonbait.md)
**Level**          |Intermediate
**User Rating**    |4.2 (25 globes from 6 users)
**Compatibility**  |C\#
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__10-5.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mandrake-dragonbait-database-connection-class__10-343/archive/master.zip)





### Source Code

```
using System;
using System.Windows.Forms; //Used for message boxes during testing
namespace Nova
{
	public class DBConnect
	{
		#region Declares
		private static ADODB.Connection cn = new ADODB.ConnectionClass();
		private ADODB.Recordset rs = new ADODB.RecordsetClass();
		private long RecordCount = 0;
		private long AbsolutePosition = 0;
		private long BMark = 0;
		#endregion
		#region Property get & let...
		public long RSCount
		{
			get{return RecordCount;}
		}
		public long AbsPosition
		{
			get{return AbsolutePosition;}
		}
		public bool EOF
		{
			get{return rs.EOF;}
		}
		public bool BOF
		{
			get{return rs.BOF;}
		}
		public long Bookmark
		{
			get{return BMark;}
			set{BMark = value;}
		}
		#endregion
		#region Methods/Functions/Procedures
		public DBConnect()
		{}//End DBConnect()
		public static void Connect(string DBName)
		{
			//Connect to the specified database. This database will
			//need to be stored in the same directory as the .exe file.
			//During design time this directory will most likely be in
			//the bin folder. So don't be confused if it doesn't find it.
			//You can also specify a longer string of subfolders for
			//DBName
			try
			{
				string DBProvider = "Provider=Microsoft.Jet.OLEDB.4.0;";
				string DBSource = "Data Source=" + Application.StartupPath.ToString() + DBName + ";";
				string ConStr = DBProvider + DBSource;
				cn.Open(ConStr, "", "", 0);
			}
			catch(Exception e)
			{
				MessageBox.Show(e.ToString());
			}
		}//End Connect()
		public void RSConnect(string TableName)
		{
			try
			{
				//Open a connection to the table
				rs.Open(TableName, cn, ADODB.CursorTypeEnum.adOpenDynamic,ADODB.LockTypeEnum.adLockOptimistic, 0);
				//Set AbsPosition and RecordCount
				if(rs.EOF == true && rs.BOF == true)
				{
					RecordCount = 0;
					AbsolutePosition = 0;
				}
				else
				{
					rs.MoveFirst();
					//Loop through records to count them
					do
					{
						RecordCount++;
						rs.MoveNext();
					} while(rs.EOF == false);
					rs.MoveFirst();
				}
			}
			catch
			{
				MessageBox.Show("Cannot connect to table specified.");
			}
		}//End RSConnect()
		public string Value(int FieldNum)
		{
			//Return a value from the database
			return(rs.Fields[FieldNum].Value.ToString());
		}//End Value()
		public void Navigate(string Direction)
		{
			//This switch statement handles the data navigation
			switch(Direction.ToUpper())
			{
				case "FIRST":
					rs.MoveFirst();
					AbsolutePosition = 1;
					break;
				case "PREVIOUS":
				case "PREV":
					if(rs.BOF == false)
					{
						rs.MovePrevious();
						AbsolutePosition--;
					}
					break;
				case "NEXT":
					if(rs.EOF == false)
					{
						rs.MoveNext();
						AbsolutePosition++;
					}
					break;
				case "LAST":
					rs.MoveLast();
					AbsolutePosition = RecordCount;
					break;
				default:
					MessageBox.Show(Direction + " is not a valid argument.");
					break;
			}//End Switch
		}//End MoveNext()
		public void GotoBookmark()
		{
			//Move to the first record
			rs.MoveFirst();
			AbsolutePosition = 1;
			//Loop through the records until the bookmark is reached
			for(long i = 1; i < BMark; i++)
			{
				rs.MoveNext();
				AbsolutePosition++;
			}
		}//End GotoBookmark
		#endregion
	}
}
```

