# xlsxdb

A tool to help inserting rows from Excel worksheet direct into a database.

<div align="center">
  <img src="./assets/xlsxdb-logo.png" alt="xlsxdb logo">
</div>


It's necessary to provide a `.env` file in the same directory of executable file, containing the connection string inside CONNECTION_STR.

Example:

```.env
CONNECTION_STR=server=xxx;user id=xxx;password=xxx;initial catalog=xxx;
```

The worksheet file should be named with the same name of destiny table.

## TODO

Generate SQL code, when table does not exist
Write logs
