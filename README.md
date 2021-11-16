# createSqliteFromExcel
creates and fills an sqlite database file from an excel file

java -jar ExcelToSqlite.jar <Excel file path (.xlsx)> [<database file path (.db)>]

 - each worksheet name will serve as table name
 - each column header should contain column-name as well as type and constrains (e.g. "col1 text primary key", "field2 varchar(10) not null" etc.)
 - the program will not check for sql syntax mistyping, so if an error occures, it will be printed to err.txt file at the current directory.
