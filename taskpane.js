// src/taskpane/taskpane.js

/**
 * Full SQL for Excel: Generate sample "menuflow" sheet, =D4ESQL formulas, and SQL code
 */
Office.onReady(() => {
  const generateBtn = document.getElementById("generate-menuflow");
  if (generateBtn) {
    generateBtn.onclick = async () => {
      await Excel.run(async (context) => {
        const sheetName = "menuflow";
        let sheet;

        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        for (let i = 0; i < sheets.items.length; i++) {
          if (sheets.items[i].name.toLowerCase() === sheetName) {
            sheets.items[i].delete();
            await context.sync();
            break;
          }
        }

        sheet = sheets.add(sheetName);

        const headers = [
          "Task Command", "Version", "Start Time", "Repetition", "Comment"
        ];
        sheet.getRange("A1:E1").values = [headers];
        sheet.getRange("A1:E1").format.fill.color = "#FCE4D6";
        sheet.getRange("A1:E1").format.font.bold = true;
        sheet.getRange("A1:E1").format.horizontalAlignment = "Left";

        const values = [
          ["getinbound user@domain.com  finance1", "", "10:00", "Every-5", "get inbound commands from this email"],
          ["sendoutbound", "", "10:00", "Every-5", "send any emails files out"],
          ["refreshtaskpane", "", "10:00", "Every-5", "refresh task pane for constant monitoring"]
        ];
        sheet.getRange("A2:E4").values = values;
        sheet.getRange("A2:E4").format.horizontalAlignment = "Left";

        sheet.getUsedRange().format.autofitColumns();
        sheet.activate();
        await context.sync();
      });
    };
  }

  const d4eBtn = document.getElementById("create-sample-d4esql");
  if (d4eBtn) {
    d4eBtn.onclick = async () => {
      await Excel.run(async (context) => {
        const sheetName = "menu.test1";
        const sheets = context.workbook.worksheets;

        sheets.load("items/name");
        await context.sync();

        for (let i = 0; i < sheets.items.length; i++) {
          if (sheets.items[i].name.toLowerCase() === sheetName.toLowerCase()) {
            sheets.items[i].delete();
            await context.sync();
            break;
          }
        }

        const sheet = sheets.add(sheetName);
        sheet.activate();

        sheet.getRange("A1:A4").values = [
          ["Rem Refresh Macro"],
          ["runsql Sales A"],
          ["runsql Sales F"],
          ["runsql Sales G"]
        ];

        const formulas = [
          { cell: "B7", formula: '=D4ESQL("Sales","A")' },
          { cell: "K7", formula: '=D4ESQL("Sales","B")' },
          { cell: "R7", formula: '=D4ESQL("Sales","C")' }
        ];

        const tips = [
          { cell: "B6", text: "Format: =D4ESQL(<SQL sheet name>,<column char>)" },
          { cell: "K6", text: "Format: =D4ESQL(<SQL sheet name>,<column char>)" },
          { cell: "R6", text: "Format: =D4ESQL(<SQL sheet name>,<column char>)" }
        ];

        for (const f of formulas) {
          sheet.getRange(f.cell).formulas = [[f.formula]];
        }

        for (const t of tips) {
          sheet.getRange(t.cell).values = [[t.text]];
        }

        await context.sync();

        const commentText = "To refresh =D4ESQL spill: Click on Get Table Data\n" +
                            "To insert new =D4ESQL: Hold CTRL key > SQL Versions > Choose SQL App option";

        for (const f of formulas) {
          const range = sheet.getRange(f.cell);
          const comment = context.workbook.comments.add(range, commentText);
          comment.author = "D4ESQL";
        }

        await context.sync();
      });
    };
  }

  const sqlBtn = document.getElementById("create-sample-sqlcode");
  if (sqlBtn) {
    sqlBtn.onclick = async () => {
      await Excel.run(async (context) => {
        const sheetName = "Sales";
        const sheets = context.workbook.worksheets;

        sheets.load("items/name");
        await context.sync();

        for (let i = 0; i < sheets.items.length; i++) {
          if (sheets.items[i].name.toLowerCase() === sheetName.toLowerCase()) {
            sheets.items[i].delete();
            await context.sync();
            break;
          }
        }

        const sheet = sheets.add(sheetName);
        sheet.activate();

/* begin SALES hardcode */

sheet.getRange("A1:A15").values = [
["/* CUS u35 */"],
["/* Driver:"],
["$type=Python,"],
["$dbms=CSV,"],
["$file_path=<<#td#>>sales_db_products.csv,"],
["$results=keep,"],
["$showresults=no,"],
["*/"],
["SELECT * FROM data WHERE Age < 45"],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("B1:B15").values = [
["/* CUSTOMER DATA */"],
["/* Driver:"],
["$type=Python,"],
["$dbms=CSV,"],
["$file_path=<<#td#>>sales_db_products.csv,"],
["$results=docs,"],
["*/"],
["SELECT * FROM data WHERE Age >= 45"],
[""],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("C1:C15").values = [
["/* Example SQL Budget Data */"],
["/* Driver: "],
["$type=Python,"],
["$dbms=Ranges,"],
["$results=docs,"],
["$rangetables=docs.BudgetData!A1:F10,docs.Sales__F!A1:F10=$rangetables,"],
["*/"],
["Select * from A"],
["where account like '5%'"],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("D1:D15").values = [
["/* Sample Oracle DB */"],
["/* Driver: "],
["$type=Python,"],
["$dbms=Oracle,"],
["$your_host=<yourhost>,"],
["$your_port=<yourport>,"],
["$your_service_name=<servicename>,"],
["$your_user=<userid>,"],
["$your_password=<password>,"],
["*/"],
["SELECT * FROM TABLE LIMIT 10"],
[""],
[""],
[""],
[""]
];
sheet.getRange("E1:E15").values = [
["/* Sample MySQL ODBC public.opendatasoft.com */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={MySQL ODBC 8.0 Unicode Driver};"],
["Server=public.opendatasoft.com;Port=3306;Database=dataset_name;User=guest;Password=guest;"],
[" =$conn,"],
["*/"],
["SELECT * FROM cities LIMIT 10;"],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("F1:F15").values = [
["/* Example SQL Budget Data */"],
["/* Driver:"],
["$type=Python,"],
["$dbms=Ranges,"],
["$results=docs,"],
["$rangetables=docs.BudgetData!A1:F10,docs.Sales__F!A1:F10=$rangetables,"],
["*/"],
["Select * from A"],
["where account like '5%'"],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("G1:G15").values = [
["/* Sample IBM DB2 */"],
["/* Driver:"],
["$type=ODBC,"],
["$conn="],
["Driver={IBM DB2 ODBC DRIVER};"],
["Database=mydatabase;"],
["Hostname=mydbhost;"],
["Port=50000;"],
["Protocol=TCPIP;"],
["Uid=myuser;"],
["Pwd=mypassword;"],
[" =$conn,"],
["*/"],
["SELECT * FROM MYTABLE"],
[""]
];
sheet.getRange("H1:H15").values = [
["/* Sample PostgreSQL ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={PostgreSQL ODBC Driver(UNICODE)};"],
["Server=mydbhost;"],
["Port=5432;"],
["Database=mydbname;"],
["Uid=myuser;"],
["Pwd=mypassword;"],
[" =$conn,"],
["*/"],
["SELECT * FROM employees LIMIT 10;"],
[""]
];
sheet.getRange("I1:I15").values = [
["/* Sample SQL Server ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={ODBC Driver 17 for SQL Server};"],
["Server=mydbhost;"],
["Database=mydbname;"],
["Uid=myuser;"],
["Pwd=mypassword;"],
[" =$conn,"],
["*/"],
["SELECT * FROM employees WHERE department = 'HR';"],
[""],
[""]
];
sheet.getRange("J1:J15").values = [
["/* Sample SQLite ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={SQLite3 ODBC Driver};"],
["Database=C:\path\to\your\database.db;"],
[" =$conn,"],
["*/"],
["SELECT * FROM customers LIMIT 10;"],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("K1:K15").values = [
["/* Sample Apache Hive ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={Cloudera ODBC Driver for Apache Hive};"],
["Host=mydbhost;"],
["Port=10000;"],
["Schema=default;"],
["Uid=myuser;"],
["Pwd=mypassword;"],
[" =$conn,"],
["*/"],
["SELECT * FROM sales_data LIMIT 10;"],
[""]
];
sheet.getRange("L1:L15").values = [
["/* Sample SAP HANA ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={HDBODBC};"],
["ServerNode=mydbhost:30015;"],
["Uid=myuser;"],
["Pwd=mypassword;"],
[" =$conn,"],
["*/"],
["SELECT * FROM sales LIMIT 10;"],
[""],
[""],
[""]
];
sheet.getRange("M1:M15").values = [
["/* Sample Amazon Redshift ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={Amazon Redshift ODBC Driver};"],
["Server=mydbhost;"],
["Port=5439;"],
["Database=mydbname;"],
["Uid=myuser;"],
["Pwd=mypassword;"],
[" =$conn,"],
["*/"],
["SELECT * FROM users LIMIT 10;"],
[""]
];
sheet.getRange("N1:N15").values = [
["/* Sample Google BigQuery ODBC Connection */"],
["/* Driver:"],
["$type=ODBC,"],
["$results=keep,"],
["$conn="],
["Driver={Simba ODBC Driver for Google BigQuery};"],
["OAuthMechanism=0;"],
["ProjectID=myprojectid;"],
["Dataset=mydataset;"],
["PrivateKeyFile=C:\path\to\your\privatekeyfile.json;"],
[" =$conn,"],
["*/"],
["SELECT * FROM `mydataset.mytable` LIMIT 10;"],
[""],
[""]
];

 





/* end SALES hardcode */



        sheet.getUsedRange().format.autofitColumns();
        sheet.getRange("A1:Z1").format.fill.color = "#FCE4D6";
        sheet.getRange("A1:Z1").format.font.bold = true;
        sheet.getRange("A1:Z1").format.horizontalAlignment = "Left";           
        await context.sync();
      });
    };
  }

/* next button here */
const helpBtn = document.getElementById("show-help-sheet");
if (helpBtn) {
  helpBtn.onclick = async () => {
    await Excel.run(async (context) => {
      const sheetName = "docs.HelpSheet";
      const sheets = context.workbook.worksheets;

      sheets.load("items/name");
      await context.sync();

      // Delete if it already exists
      for (let i = 0; i < sheets.items.length; i++) {
        if (sheets.items[i].name.toLowerCase() === sheetName.toLowerCase()) {
          sheets.items[i].delete();
          await context.sync();
          break;
        }
      }

      const sheet = sheets.add(sheetName);
      sheet.activate();

      // Add placeholder content in Column A
sheet.getRange("A1:A15").values = [
["Button"],
["Add SQL File"],
["Get Data Table"],
["Get Menu"],
["Remove Sheets***"],
["Show/Hide SQL Sheets"],
["Show/Hide Docs"],
["Format SQL"],
["Clear Addin"],
["About D4E"],
["Run SQL"],
["SQL Versions"],
["Run Macros"],
[""],
[""]
];
sheet.getRange("B1:B15").values = [
["Meaning"],
["Open a file containing SQL codes and include in D4E App Store"],
["Run SQL attached to the Sheet and Column Name"],
["Read all SQL sheets and generate menu for each SQL code"],
["Delete sheets with name beginning with 'Sheets'"],
["Show or Hide sheets that do not begin wth 'menu', 'docs', 'macro','sheet'"],
["Show or Hide sheets beginning with 'docs'"],
["Color format all SQL sheets"],
["Clear contents of D4E Store"],
["Show D4E Info and Task Pane"],
["Show Menu of SQL (default Version)"],
["Show Menu of SQL  with ALL Versions"],
["Show Menu of ALL Macros"],
[""],
[""]
];
sheet.getRange("C1:C15").values = [
["Button"],
["Start Scheduler"],
["Stop Scheduler"],
["Show Task Status"],
["Validate Tasks"],
["Record Macro while Executing"],
["Record Macro only (No Exec)"],
["Stop Macro recording"],
["Run this Macro"],
["Show Param List"],
["Clear Param List"],
["Update Table"],
[""],
[""],
[""]
];
sheet.getRange("D1:D15").values = [
["Meaning"],
["Start the Scheduler and schedule any tasks found"],
["Stop the Scheduler and clear outstanding tasks"],
["Show the tasks scheduled - when to execute next"],
["Validate tasks in menuflow"],
["Record the choices made in Run SQL, Run SQL Versions, Run Macros while executing them"],
["Record the choices made in Run SQL, Run SQL Versions, Run Macros only - do not execute any SQL"],
["Stop recording choices in menu"],
["Run the macro commands found in Column A of this sheet"],
["Show the placeholder values to use for the parameters"],
["Clear the internal parameter list set up by setparam command."],
["Read Action Code column and update (U), delete (D) or insert (I) the row to the database defined by the sheet name"],
[""],
[""],
[""]
];
sheet.getRange("E1:E15").values = [
["|"],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("F1:F15").values = [
["Macro Commands"],
["about - Display about information, activate/refresh Excel Taskpane."],
["aboutinfo - see about."],
["add1sqlfile - see addsqlfile."],
["addsqlfile <fullfilename> - Add an SQL file to the D4E store and activate RUN Mode."],
["cleanup - Reset status of the add-in."],
["clearaddin - Clear the data in add-in or store."],
["clearlog - Clear contents of log (menulog)."],
["clearparamlist - Clear parameters in param_list."],
["copyto <sheet> - Copy last SQL output to <sheet> which is used by a pivot table or report."],
["createsendfile <email> <title> <file> <maxlines> - Prepare output results for sending by email."],
["createsendsheet <email> <title> <sheet> <+/- maxlines> - Prepare report sheet for sending by email."],
["encode <text> - Encode a value using a D4E encoding scheme."],
["excelvba <file!VBAmacro> - Run VBA code."],
[""]
];
sheet.getRange("G1:G15").values = [
["Macro Commands"],
["formatsql - Paint or format all SQL code in sheets."],
["get - see runsql."],
["getdatatable - see runsql."],
["getdatatablev - see runsqlv."],
["getinbound <email|dom> <topic> - Read commands from user@domain.com or @domain.com for topic"],
["getmenu - Create menu sheet and menu commandbars."],
["help <cmd|searchkey> - Display help on cmd or searchkey."],
["initmenuflow - Make menuflow from SQL sheets."],
["logmsg <msg> - Log a message."],
["openfile - see opensqlfile."],
["opensqlfile <file> - Open SQL file for EDIT Mode."],
["prepareupdate - Prepare activesheet to have Action Code column."],
["quiet_mode - see silent_mode."],
[""]
];
sheet.getRange("H1:H15").values = [
["Macro Commands"],
["quietmode - see silent_mode."],
["reademailmacro - see getinbound."],
["refreshpivots <sheet> - Refresh pivot tables or reports."],
["rem - This is a comment line."],
["removesheets - Delete output sheets (Sheets???)."],
["runmacro <sheet> - Run macro command from sheet."],
["runquiet - see silent_mode."],
["runsilent - see silent_mode."],
["runsql <sheet> <colummABC> - Execute an SQL query found in column X of sheet."],
["runsqlkeep <sheet> <columnABC> - Execute a runsql command but keeps the results."],
["runsqlv <sheet> <version_title> - Execute SQL in column X, with version_title in row 1."],
["runsqlvkeep <sheet> <version_title> - Execute a runsqlv command but keeps the results."],
["sendoutbound - Send output results to email."],
[""]
];
sheet.getRange("I1:I15").values = [
["Macro Commands"],
["setdir - see settempdir."],
["setfolder - see settempdir."],
["setlogrows <maxcount> - Set max log entries to display."],
["setparam <param>=<value> - Set a value for the placeholder and add to param_list."],
["setprompt - see setparam."],
["settempdir <fullfoldername> - Set or create the folder name for temporary output files."],
["settempfolder - see settempdir."],
["shell <cmd> - Execute a shell command."],
["showdocs - Show sheet named 'docs.*' ."],
["silent_mode <ON|OFF> - ON = Do not prompt user, use pre-set param_list instead."],
["silentmode - see silent_mode."],
["startscheduler - Start the scheduler."],
["stopscheduler - Stop the scheduler."],
[""]
];
sheet.getRange("J1:J15").values = [
["Macro Commands"],
["update - see updatedb."],
["updatedb <sheet> - Update database table from sheet rows with Action code."],
["updatedbfile <filename> - Update database table from pathfilename rows with Action code."],
["updatefile - see updatedbfile."],
["validatetasks - Validate all tasks."],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("K1:K15").values = [
["|"],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("L1:L15").values = [
["Driver Variables"],
["$conn=<value>=$conn, - Indicates ODBC connection string"],
["$dbms=<value>, - DBMS type to process Python, ODBC, Excel, Ranges, Sheets, CSV"],
["$file_path=<value>, - Full file path"],
["$file=<value>, - File name"],
["$filefolder=<value>, - Folder and File name"],
["$filetables=<value>, - Folder and File name"],
["$pythondebug=<value>, - Turn ON/OFF Python debug messages"],
["$rangeheaders=<value>, - Header names"],
["$rangetables=<value>=$rangetables, - Indicates ranges as tables (A,B,C,etc)"],
["$results=<value>, - Keep= Save results to file, Docs=Save results as docs sheets"],
["$sheettables=<value>, - List of sheets to load as tables, otheres are skipped"],
["$showresults=<value>, - No- do nto creat sheet, else create output sheet"],
["$sql_query=<value>=$sql_query, - Default SQL Query if none found"],
[""]
];
sheet.getRange("M1:M15").values = [
["Driver Variables"],
["$table=<value>, - Table to be loaded"],
["$tables=<value>=$tables, - Table to be loaded"],
["$type=<value>, - ODBC, Python, Excel, CSV"],
["$workbook=<value>, - Source workbook of sheets for table loading, default Activeworkbook"],
["$your_database=<value>, - Python  parameter for accessing DBMS"],
["$your_host=<value>, - Python  parameter for accessing DBMS"],
["$your_password=<value>, - Python  parameter for accessing DBMS"],
["$your_port=<value>, - Python  parameter for accessing DBMS"],
["$your_server=<value>, - Python  parameter for accessing DBMS"],
["$your_service_name=<value>, - Python  parameter for accessing DBMS"],
["$your_user=<value>, - Python  parameter for accessing DBMS"],
[""],
[""],
[""]
];
      // End of placeholder content in Column A      

      sheet.getUsedRange().format.autofitColumns();
        sheet.getRange("A1:Z1").format.fill.color = "#FCE4D6";
        sheet.getRange("A1:Z1").format.font.bold = true;
        sheet.getRange("A1:Z1").format.horizontalAlignment = "Left";      
      await context.sync();
      console.log("Help sheet created.");
    });
  };
}


/* end of button */

/* next button below */
  const budgetBtn = document.getElementById("create-budget-app");
  if (budgetBtn) {
    budgetBtn.onclick = async () => {
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;

        // Delete if existing
        const sheetNames = ["docs.FIS.Data", "menu.FIS.D4ESQL", "FISSQL"];
        sheets.load("items/name");
        await context.sync();

        for (const name of sheetNames) {
          for (let i = 0; i < sheets.items.length; i++) {
            if (sheets.items[i].name.toLowerCase() === name.toLowerCase()) {
              sheets.items[i].delete();
              await context.sync();
              break;
            }
          }
        }

        // Create sheet docs.FIS.Data
        let sheet = sheets.add("docs.FIS.Data");
// start data placeholder        
sheet.getRange("A1:A15").values = [
["Budget_Code"],
["ActualSYD_202501"],
["ActualSYD_202502"],
["ActualSYD_202503"],
["ActualSYD_202504"],
["ActualNOU_202501"],
["ActualNOU_202502"],
["ActualNOU_202503"],
["ActualNOU_202504"],
["ActualPRA_202501"],
["ActualPRA_202502"],
["ActualPRA_202503"],
["ActualPRA_202504"],
[""],
[""]
];
sheet.getRange("B1:B15").values = [
["Company"],
["10"],
["10"],
["10"],
["10"],
["20"],
["20"],
["20"],
["20"],
["30"],
["30"],
["30"],
["30"],
[""],
[""]
];
sheet.getRange("C1:C15").values = [
["Branch"],
["Sydney"],
["Sydney"],
["Sydney"],
["Sydney"],
["Noumea"],
["Noumea"],
["Noumea"],
["Noumea"],
["Prague"],
["Prague"],
["Prague"],
["Prague"],
[""],
[""]
];
sheet.getRange("D1:D15").values = [
["Account"],
["400303"],
["400101"],
["500404"],
["500405"],
["400303"],
["400101"],
["500404"],
["500405"],
["400303"],
["400101"],
["500404"],
["500405"],
[""],
[""]
];
sheet.getRange("E1:E15").values = [
["Curr"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
["AUD"],
[""],
[""]
];
sheet.getRange("F1:F15").values = [
["Net_Amt"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
["800"],
[""],
[""]
];
//end data placeholder

      sheet.getUsedRange().format.autofitColumns();
        sheet.getRange("A1:Z1").format.fill.color = "#FCE4D6";
        sheet.getRange("A1:Z1").format.font.bold = true;
        sheet.getRange("A1:Z1").format.horizontalAlignment = "Left";    

        // Create sheet menu.FIS.D4ESQL
        sheet = sheets.add("menu.FIS.D4ESQL");
        // place holder
sheet.getRange("A1:A3").values = [
["Rem Refresh Macro"],
["runsql FISSQL A"],
["runsql FISSQL B"]
];

sheet.getRange("Q1:Q15").values = [
["Readme"],
["This is the Sample FIS Budget Application that will not need external data. It will demonstrate use of Full SQL inside Excel via =D4ESQL function."],
["SHEETS CREATED ARE:"],
["1. docs.FIS.DATA - contains sample data and will be read as a table in SQL"],
["2. FISSQL - contains SQL code to read Excel sheets specifically docs.FIS.DATA"],
["3. menu.FIS.D4ESQL - has =D4ESQL formula to refer to SQL code to report on the table containing Budget data."],
[""],
["TO RUN OR REFRESH:"],
["From anywhere in workbook, hold SHIFT then click on Get Data Table. This will refresh all cells with =D4ESQL."],
["TO SEE IT WORKING:"],
["1. Initially, SQL reports in 'menu.FIS.D4ESQL' shows $800 as net amounts."],
["2. Click on Get Menu to refresh the menus. Goto SQL Versions > FISSQL > FIS Budget Data Update > this will create an update sheet."],
["3. Change amounts in Budget Update sheet, eg $500, then click Update Table.  Internal table (docs.FIS.DATA) will be changed. "],
["4. Hold SHIFT then click on Get Data Table and SQL reports will reflect new amounts."],
[""]
];

        const formulas = [
          { cell: "B7", formula: '=D4ESQL("FISSQL","A")' },
          { cell: "I7", formula: '=D4ESQL("FISSQL","B")' }
        ];

        const tips = [
          { cell: "B6", text: "BUDGETS FOR GROUP FIVE ACCOUNT" },
          { cell: "I6", text: "BUDGETS FOR GROUP FOUR ACCOUNT" }
        ];

        for (const f of formulas) {
          sheet.getRange(f.cell).formulas = [[f.formula]];
        }

        for (const t of tips) {
          sheet.getRange(t.cell).values = [[t.text]];
        }

        await context.sync();

        const commentText = "To refresh =D4ESQL spill: Click on Get Data Table\n" +
                            "To insert new =D4ESQL: Hold CTRL key > SQL Versions > Choose SQL App option";

        for (const f of formulas) {
          const range = sheet.getRange(f.cell);
          const comment = context.workbook.comments.add(range, commentText);
          comment.author = "D4ESQL";
        }        

        // end placeholder
      sheet.getUsedRange().format.autofitColumns();
        sheet.getRange("A1:Z1").format.fill.color = "#FCE4D6";
        sheet.getRange("A1:Z1").format.font.bold = true;
        sheet.getRange("A1:Z1").format.horizontalAlignment = "Left";    


        // Create sheet FIS.SQL
        sheet = sheets.add("FISSQL");
        // data place holder
sheet.getRange("A1:A15").values = [
["/* FIS Budget Data 5 */"],
["/* Driver: "],
["$type=Python,"],
["$dbms=Ranges,"],
["$results=docs,"],
["$tables=docs.FIS.Data!A1:F15,docs.FIS.DATA!A:F=$tables,"],
["*/"],
["Select * from A"],
["where account like '5%'"],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("B1:B15").values = [
["/* FIS Budget Data 4 */"],
["/* Driver: "],
["$type=Python,"],
["$dbms=Ranges,"],
["$results=docs,"],
["$tables=docs.FIS.Data!A1:F15,docs.FIS.DATA!A:F=$tables,"],
["*/"],
["Select * from B"],
["where account like '4%'"],
[""],
[""],
[""],
[""],
[""],
[""]
];
sheet.getRange("C1:C15").values = [
["/* FIS Budget Data Update */"],
["/* Driver: "],
["$type=Python,"],
["$dbms=Ranges,"],
[""],
["$tables=docs.FIS.Data!A1:F15,docs.FIS.DATA!A:F=$tables,"],
["*/"],
["SELECT "],
["   'U' as Action_Code,"],
["Budget_Code as Budget_Code_key,Company,Branch,Account,Curr,Net_Amt"],
["FROM B"],
["where 1=1"],
["and Branch = 'Sydney'"],
[""],
[""]
];
        // end data placeholder

      sheet.getUsedRange().format.autofitColumns();
        sheet.getRange("A1:Z1").format.fill.color = "#FCE4D6";
        sheet.getRange("A1:Z1").format.font.bold = true;
        sheet.getRange("A1:Z1").format.horizontalAlignment = "Left";    

        await context.sync();
        console.log("Budget app sheets created.");
      });
    };
  }


/* end of button */

/* next button below */

/* end of button */

});
