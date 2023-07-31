using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;


public class Program
{

    int tableNum = 0;
    Dictionary<string, int> tableNameDict = new Dictionary<string, int>();
    const string metaDataWorkSheet = "MetaData";

    public static void Main(string[] args)
    {
        string dbLeftOutput;
        string dbRightOutput;
        string? dbLeftInput = null;
        string? dbRightInput = null;
        string option;  //1. Input 2 DBs on the cli, 2. Input 2 generated excel workbooks on the cli, 3. Input 2 DBs using the args, 4. Input 2 generated excel wbs using args
        string? exportLocation = null;
        string? tempOption = null;

        var tableNameDict = new Dictionary<string, int>();
        var tableNum = 0;
        Console.WriteLine("Compare db on left with right");
        if (args.Length == 4)
        {
            tempOption = args[0];
            dbLeftInput = args[1];
            dbRightInput = args[2];
            exportLocation = args[3];
            if (exportLocation.EndsWith("\""))
            {
                exportLocation = args[3].Substring(0, args[3].Length - 1);
            }
        }

        Console.WriteLine("Select \n1. 2 dbs\n2. 2 generated excel workbook ?");
        if (tempOption != null)
            option = tempOption.Trim();
        else
            option = Console.ReadLine()!.Trim();
        var timeNow = ((long)(DateTime.Now - new DateTime(1970, 1, 1)).TotalSeconds).ToString();
        if (option == "1")
        {
            Console.WriteLine("Left DB location: ");

            dbLeftInput = Console.ReadLine()!;
            Console.WriteLine("\nDB Right Location: ");
            dbRightInput = Console.ReadLine()!;
            Console.WriteLine("Where to store the resulting files ?");
            exportLocation = Console.ReadLine()!;

            if (!checkDirs(dbLeftInput, dbRightInput)) throw new Exception("Either of the db files doesn't exist");
            checkExportLocation(exportLocation);

            dbLeftOutput = Path.Combine(exportLocation, "dbLeft_" + timeNow + ".xlsx");
            dbRightOutput = Path.Combine(exportLocation, "dbRight_" + timeNow + ".xlsx");

            exportDB(dbLeftInput, dbLeftOutput, ref tableNum, ref tableNameDict);
            exportDB(dbRightInput, dbRightOutput, ref tableNum, ref tableNameDict);
        }
        else if (option == "2")
        {
            Console.WriteLine("Left Excel location: ");

            dbLeftInput = Console.ReadLine()!;
            Console.WriteLine("\nExcel Right Location: ");
            dbRightInput = Console.ReadLine()!;
            Console.WriteLine("Where to store the resulting files ?");
            exportLocation = Console.ReadLine()!;

            if (!checkDirs(dbLeftInput, dbRightInput)) throw new Exception("Either of the excel files doesn't exist");

            dbLeftOutput = dbLeftInput;
            dbRightOutput = dbRightInput;
        }
        else if (option == "3" && dbLeftInput != null && dbRightInput != null && exportLocation != null)
        {
            Console.WriteLine($"Left DB location: {dbLeftInput}");

            Console.WriteLine($"\nDB Right Location: {dbRightInput}");
            Console.WriteLine($"Resulting files: {exportLocation}");

            if (!checkDirs(dbLeftInput, dbRightInput)) throw new Exception("Either of the db files doesn't exist");

            dbLeftOutput = Path.Combine(exportLocation, "dbLeft_" + timeNow + ".xlsx");
            dbRightOutput = Path.Combine(exportLocation, "dbRight_" + timeNow + ".xlsx");

            exportDB(dbLeftInput, dbLeftOutput, ref tableNum, ref tableNameDict);
            exportDB(dbRightInput, dbRightOutput, ref tableNum, ref tableNameDict);
        }
        else if (option == "4" && dbLeftInput != null && dbRightInput != null && exportLocation != null)
        {
            Console.WriteLine($"Left Excel location: {dbLeftInput}");

            Console.WriteLine($"\nExcel Right Location: {dbRightInput}");
            Console.WriteLine($"Resulting files: {exportLocation}");

            dbLeftOutput = dbLeftInput;
            dbRightOutput = dbRightInput;
        }
        else throw new Exception("Unknown option");

        if (exportLocation == null)
            throw new Exception("Unknown export location");
        try
        {
            checkExportLocation(exportLocation!);
            var mergeFile = Path.Combine(exportLocation!, "MergeResult_" + timeNow + ".xlsx");

            generateMergeFile(dbLeftOutput, dbRightOutput, mergeFile);
            var sInfo = new ProcessStartInfo(mergeFile);
            sInfo.Verb = "Open";
            sInfo.UseShellExecute = true;
            Process.Start(sInfo);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Export & Merge failed because {ex.Message}");
        }
    }

    private static void checkExportLocation(string exportLocation)
    {
        var dir = new DirectoryInfo(exportLocation);
        if (!dir.Exists)
            dir.Create();
    }

    private static bool checkDirs(string dbLeft, string dbRight)
    {
        return File.Exists(dbLeft) && File.Exists(dbRight);
    }

    public static void exportDB(string dbLocation, string exportPath, ref int tableNum, ref Dictionary<string, int> tableNameDict)
    {
        Console.WriteLine($"Exporting {dbLocation}");
        var localTableNameDict = new Dictionary<string, int>();

        if (!OperatingSystem.IsWindows()) throw new Exception("Only windows supported!");

        string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbLocation;
        using (var cn = new OleDbConnection(connectionString))
        {
            cn.Open();

            if (File.Exists(exportPath))
            {
                throw new Exception("db File already exists");
            }
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet(metaDataWorkSheet);        //assumptuion: index 0
            using (FileStream fs = new FileStream(exportPath, FileMode.Create, FileAccess.ReadWrite))
            {
                //Get all tables in the db
                DataTable tables = cn.GetSchema("Tables");




                var tableNames = new List<String>();

                //Get all table names
                foreach (DataRow row in tables.Rows)
                {
                    var tableName = row[2].ToString()!;

                    //Only take table names that start with tbl
                    if (tableName.ToLower().StartsWith("tbl"))
                        tableNames.Add(tableName);
                }



                foreach (var tableName in tableNames)
                {
                    Console.WriteLine($"Exporting {tableName} of the db");
                    FillErrorEventHandler errorEventHandler = delegate (object sender, FillErrorEventArgs args)
                    {
                        //Console.WriteLine($"{tableName} has a column with invalid values. The error is: {args?.Errors?.Message}");
                        if (args != null)
                        {
                            Console.WriteLine($"{tableName} has a row with invalid values. The error is: {args?.Errors?.Message} and the row id is {args!.Values[0]?.ToString()}");

                            //args?.DataTable.Rows.Add(row);
                            args!.Continue = true;
                        }
                    };

                    /*OleDbCommand cmd = cn.CreateCommand();
                    //cmd.Transaction = cn.BeginTransaction();
                    cmd.CommandText = "Select * from tblRefAccountingChange";*/

                    using (OleDbDataReader reader = new OleDbCommand("Select * from " + tableName, cn).ExecuteReader())
                    {
                        //Create worksheet in the excel with the tablename and an identifier
                        var sheetName = tableName;
                        if (sheetName.Length == 0) throw new Exception("Sheet name is empty!");
                        //BUG: Some tables have >31 chars in name which isn't allowed by Excel, instead we simply map unique table names to unique ints and assign these ints as sheet names

                        if (!tableNameDict.ContainsKey(sheetName))
                        {
                            tableNameDict[sheetName] = ++tableNum;

                        }
                        localTableNameDict[sheetName] = tableNameDict[sheetName];

                        sheetName = tableNameDict[sheetName].ToString();

                        var sheet = workbook.CreateSheet(sheetName);

                        var tableSchema = reader.GetSchemaTable()!;

                        /* var row0 = sheet.CreateRow(0);
                         for (int i = 0; i < tableSchema.Columns.Count; ++i)
                         {
                             DataColumn column = tableSchema.Columns[i];
                             var cell = row0.CreateCell(i);
                             cell.SetCellValue(column.ColumnName);
                         }*/

                        //Set first row values to be column names of the table
                        var row0 = sheet.CreateRow(0);

                        for (var rowNum = 0; rowNum < tableSchema.Rows.Count; ++rowNum)
                        {
                            var cell = row0.CreateCell(rowNum);
                            DataRow dataRow = tableSchema.Rows[rowNum];

                            cell.SetCellValue(dataRow[tableSchema.Columns[0]].ToString());
                        }
                        var dt = new DataTable();
                        dt.Load(reader, LoadOption.PreserveChanges, errorEventHandler);
                        Console.WriteLine($"Exported row:");
                        reader.Close();

                        //Now iterate over each row and store its values in the excel.
                        var rowNums = 1;
                        foreach (DataRow rowR in dt.Rows)
                        {
                            var row = sheet.CreateRow(rowNums);
                            for (int colNum = 0; colNum < dt.Columns.Count; ++colNum)
                            {
                                var cell = row.CreateCell(colNum);
                                cell.SetCellValue(rowR[colNum].ToString());
                            }
                            Console.Write($"\r{rowNums} / {dt.Rows.Count}");
                            rowNums++;
                        }


                        Console.WriteLine($"\n \n \n Exported {tableName}");

                    }
                }

                createMetaData(ref workbook, ref localTableNameDict);



                workbook.Write(fs);
                Console.WriteLine($"Exported db successfully at {exportPath}");
            }
            cn.Close();
        }
    }

    private static void createMetaData(ref XSSFWorkbook workbook, ref Dictionary<string, int> tableNameDict)
    {
        var sheet = workbook.GetSheet(metaDataWorkSheet);
        var rowCount = 1;
        var row = sheet.CreateRow(rowCount++);
        var cell = row.CreateCell(0);
        cell.SetCellValue("MetaData");
        row = sheet.CreateRow(rowCount++);
        cell = row.CreateCell(0);
        cell.SetCellValue("Table Name");
        cell = row.CreateCell(1);
        cell.SetCellValue("Unique Int Value");

        for (int i = 0; i < tableNameDict.Count(); ++i)
        {

            var localRow = sheet.CreateRow(rowCount + i);
            localRow.CreateCell(0).SetCellValue(tableNameDict.ElementAt(i).Key);
            localRow.CreateCell(1).SetCellValue(tableNameDict.ElementAt(i).Value.ToString());
        }
    }

    //TODO: Break function into smaller parts to allow GC to sweep unused objects
    //TODO: Optimize to use more in-disk objects with greater read time than  in-memory objects with greater memory size
    public static void generateMergeFile(string dbLeftFile, string dbRightFile, string mergeFile)
    {
        Console.WriteLine($"Generating Merged DB Workbook {mergeFile}");

        if (!OperatingSystem.IsWindows()) throw new Exception("Only windows supported!");
        if (File.Exists(mergeFile))
        {
            throw new Exception("Merge File already exists");
        }
        var workbook = new XSSFWorkbook();
        var tableDict = new Dictionary<int, string>(); //can simply use a list of size n, but its not a guarantee each ID will be within the range [1,n] where n is no. of tables.
        using (FileStream fs = new FileStream(mergeFile, FileMode.Create, FileAccess.ReadWrite))
        {

            var sheet0 = workbook.CreateSheet("MetaData");
            int rowCount = 0;
            createMetaDataWorksheet(ref sheet0, ref rowCount);
            var redCellStyle = createCellStyle(ref workbook, NPOI.HSSF.Util.HSSFColor.Red.Index);
            var yellowCellStyle = createCellStyle(ref workbook, NPOI.HSSF.Util.HSSFColor.Yellow.Index);
            var greenCellStyle = createCellStyle(ref workbook, NPOI.HSSF.Util.HSSFColor.Green.Index);
            var orangeCellStyle = createCellStyle(ref workbook, NPOI.HSSF.Util.HSSFColor.Orange.Index);
            var blueCellStyle = createCellStyle(ref workbook, NPOI.HSSF.Util.HSSFColor.Blue.Index);
            var cyanCellStyle = createCellStyle(ref workbook, NPOI.HSSF.Util.HSSFColor.SkyBlue.Index);

            //main logic
            using (FileStream leftfs = new FileStream(dbLeftFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var workbookLeft = new XSSFWorkbook(leftfs)!;
                using (FileStream rightfs = new FileStream(dbRightFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var workbookRight = new XSSFWorkbook(rightfs)!;

                    var outRight = new List<string>(); // Tables that are not in right
                    var outLeft = new List<string>(); // Tables that are not in left
                    var dictWorkbookLeft = new Dictionary<String, int>();
                    var dictWorkbookRight = new Dictionary<String, int>();
                    var hashleft = new HashSet<String>();
                    var hashright = new HashSet<String>();

                    tableDict = getTableMapping(ref workbookLeft, ref workbookRight);


                    Console.WriteLine($"MergedWorkbook: Creating Metadata Worksheet");

                    for (var sheetNum = 0; sheetNum < workbookLeft.NumberOfSheets; ++sheetNum)
                    {
                        if (workbookLeft.GetSheetName(sheetNum) != metaDataWorkSheet)
                        {
                            dictWorkbookLeft.Add(workbookLeft.GetSheetName(sheetNum), sheetNum);
                            hashleft.Add(workbookLeft.GetSheetName(sheetNum));
                        }
                    }

                    for (var sheetNum = 0; sheetNum < workbookRight.NumberOfSheets; ++sheetNum)
                    {
                        if (workbookRight.GetSheetName(sheetNum) != metaDataWorkSheet)
                        {
                            dictWorkbookRight.Add(workbookRight.GetSheetName(sheetNum), sheetNum);
                            hashright.Add(workbookRight.GetSheetName(sheetNum));

                        }
                    }

                    outRight = hashleft.Except(hashright).ToList(); //Removes all elements that are in right & left from the left
                    outLeft = hashright.Except(hashleft).ToList();

                    hashleft.IntersectWith(hashright); //Makes the object to contain elements only if they are present in both objects

                    Console.WriteLine($"MergedWorkbook: Checking tables of both dbs");

                    rowCount++; //Row to start inserting table metadata from
                    foreach (var item in hashleft)
                    {
                        int intItem = int.Parse(item);
                        Console.WriteLine($"MergedWorkbook: Working on {item}: {tableDict[intItem]}");
                        var indexLeft = dictWorkbookLeft[item];
                        var indexRight = dictWorkbookRight[item];

                        var sheet = workbook.CreateSheet(item);
                        var sheetLeft = workbookLeft.GetSheetAt(indexLeft);
                        var sheetRight = workbookRight.GetSheetAt(indexRight);

                        var row0 = sheet0.CreateRow(rowCount);
                        row0.CreateCell(0).SetCellValue(item);
                        row0.CreateCell(1).SetCellValue(tableDict[intItem]);

                        int leftRowCount = sheetLeft.LastRowNum + 1;
                        int rightRowCount = sheetRight.LastRowNum + 1;
                        int leftColumnCount = leftRowCount - 1 > 0 ? sheetLeft.GetRow(0).LastCellNum : 0;
                        int rightColumnCount = rightRowCount - 1 > 0 ? sheetRight.GetRow(0).LastCellNum : 0;

                        var dictLeftRows = new Dictionary<string, int>();
                        var dictRightRows = new Dictionary<string, int>();
                        var dictLeftRowsInverse = new Dictionary<int, string>();
                        var dictRightRowsInverse = new Dictionary<int, string>();
                        var leftRows = new HashSet<String>();
                        var rightRows = new HashSet<String>();

                        var dictLeftCols = new Dictionary<string, int>();
                        var dictRightCols = new Dictionary<string, int>();
                        var dictLeftColsInverse = new Dictionary<int, string>();
                        var dictRightColsInverse = new Dictionary<int, string>();
                        var leftCols = new HashSet<String>();
                        var rightCols = new HashSet<String>();

                        var reds = new HashSet<String>();
                        var yellows = new HashSet<String>();
                        var greens = new HashSet<String>();
                        var oranges = new HashSet<String>();
                        var blues = new HashSet<String>();
                        var cyans = new HashSet<String>();

                        Console.WriteLine($"MergedWorkbook: Getting all headers for the table from left");
                        //Getting headers for every column
                        if (leftRowCount > 0)
                        {
                            for (int i = 0; i < leftColumnCount; ++i)
                            {
                                leftCols.Add(sheetLeft.GetRow(0).GetCell(i).StringCellValue);
                                dictLeftCols.Add(sheetLeft.GetRow(0).GetCell(i).StringCellValue, i);
                                dictLeftColsInverse.Add(i, sheetLeft.GetRow(0).GetCell(i).StringCellValue);
                            }
                        }

                        Console.WriteLine($"MergedWorkbook: Getting all headers for the table from right");
                        if (rightRowCount > 0)
                        {
                            for (int i = 0; i < rightColumnCount; ++i)
                            {
                                rightCols.Add(sheetRight.GetRow(0).GetCell(i).StringCellValue);
                                dictRightCols.Add(sheetRight.GetRow(0).GetCell(i).StringCellValue, i);
                                dictRightColsInverse.Add(i, sheetRight.GetRow(0).GetCell(i).StringCellValue);
                            }
                        }

                        Console.WriteLine($"MergedWorkbook: Getting all IDs for the table from left");
                        //Getting IDs for every row
                        if (leftColumnCount > 0)
                        {
                            for (int i = 1; i < leftRowCount; ++i)
                            {
                                leftRows.Add(sheetLeft.GetRow(i).GetCell(0).StringCellValue);
                                dictLeftRows.Add(sheetLeft.GetRow(i).GetCell(0).StringCellValue, i);
                                dictLeftRowsInverse.Add(i, sheetLeft.GetRow(i).GetCell(0).StringCellValue);
                            }
                        }
                        Console.WriteLine($"MergedWorkbook: Getting all IDs for the table from right");
                        if (rightColumnCount > 0)
                        {
                            for (int i = 1; i < rightRowCount; ++i)
                            {
                                rightRows.Add(sheetRight.GetRow(i).GetCell(0).StringCellValue);
                                dictRightRows.Add(sheetRight.GetRow(i).GetCell(0).StringCellValue, i);
                                dictRightRowsInverse.Add(i, sheetRight.GetRow(i).GetCell(0).StringCellValue);
                            }
                        }

                        reds = leftRows.Except(rightRows).ToHashSet<string>();
                        yellows = rightRows.Except(leftRows).ToHashSet<string>();
                        blues = leftCols.Except(rightCols).ToHashSet<string>();
                        cyans = rightCols.Except(leftCols).ToHashSet<string>();

                        var dictCols = new Dictionary<string, int>();
                        var dictRows = new Dictionary<string, int>();
                        var dictColsInverse = new Dictionary<int, string>();
                        var dictRowsInverse = new Dictionary<int, string>();

                        var allCols = leftCols.Union(rightCols).ToHashSet();
                        var allRows = leftRows.Union(rightRows).ToHashSet();
                        int columns = allCols.Count;
                        int rows = allRows.Count;

                        Console.WriteLine($"MergedWorkbook: Creating headers in merged worksheet");
                        //Creating ID column in merged worksheet
                        if (rows > 0)
                        {
                            var row = sheet.CreateRow(0);

                            var tempColCount = 0;
                            foreach (var elem in allCols)
                            {
                                row.CreateCell(tempColCount).SetCellValue(elem);
                                dictCols.Add(elem, tempColCount);
                                dictColsInverse.Add(tempColCount, elem);
                                tempColCount++;
                            }
                        }
                        Console.WriteLine($"MergedWorkbook: Creating IDs in merged worksheet");
                        //Creating Header row in the merged worksheet
                        if (columns > 0)
                        {
                            //Row is >0 is a given.

                            var tempRowCount = 0;
                            foreach (var elem in allRows)
                            {
                                var row = sheet.CreateRow(tempRowCount + 1);
                                row.CreateCell(0).SetCellValue(elem);
                                dictRows.Add(elem, tempRowCount);
                                dictRowsInverse.Add(tempRowCount, elem);
                                tempRowCount++;
                            }
                        }

                        var tempColor = NPOI.HSSF.Util.HSSFColor.White.Index;
                        var rowColors = Enumerable.Range(0, rows).Select(i => tempColor).ToList();
                        var colColors = Enumerable.Range(0, columns).Select(i => tempColor).ToList();

                        Console.WriteLine($"MergedWorkbook: Merging Rows&Cols in Worksheet");
                        for (int i = 1; i <= rows; ++i)
                        {
                            var rowVal = dictRowsInverse[i - 1];
                            var row = sheet.GetRow(i);
                            for (int j = 0; j < columns; ++j)
                            {
                                var colVal = dictColsInverse[j];

                                var cell = row.CreateCell(j);

                                if (leftCols.Contains(colVal) && rightCols.Contains(colVal))
                                {
                                    if (leftRows.Contains(rowVal) && rightRows.Contains(rowVal))
                                    {
                                        var leftVal = sheetLeft.GetRow(dictLeftRows[rowVal]).GetCell(dictLeftCols[colVal]).StringCellValue;
                                        var rightVal = sheetRight.GetRow(dictRightRows[rowVal]).GetCell(dictRightCols[colVal]).StringCellValue;

                                        if (leftVal == rightVal)
                                        {
                                            cell.SetCellValue(leftVal);
                                        }
                                        else
                                        {
                                            cell.SetCellValue("Left: " + leftVal + " Right: " + rightVal);
                                            rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Orange.Index;

                                            oranges.Add(rowVal); //hashset cant contain duplicates so this is ok
                                        }
                                    }
                                    else if (leftRows.Contains(rowVal))
                                    {
                                        cell.SetCellValue(sheetLeft.GetRow(dictLeftRows[rowVal]).GetCell(dictLeftCols[colVal]).StringCellValue);
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Red.Index;
                                    }
                                    else if (rightRows.Contains(rowVal))
                                    {
                                        cell.SetCellValue(sheetRight.GetRow(dictRightRows[rowVal]).GetCell(dictRightCols[colVal]).StringCellValue);
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                                    }
                                }
                                else if (leftCols.Contains(colVal))
                                {
                                    if (leftRows.Contains(rowVal) && rightRows.Contains(rowVal))
                                    {
                                        cell.SetCellValue(sheetLeft.GetRow(dictLeftRows[rowVal]).GetCell(dictLeftCols[colVal]).StringCellValue);
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Orange.Index;

                                        oranges.Add(rowVal);
                                    }
                                    else if (leftRows.Contains(rowVal))
                                    {
                                        cell.SetCellValue(sheetLeft.GetRow(dictLeftRows[rowVal]).GetCell(dictLeftCols[colVal]).StringCellValue);
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Red.Index;
                                    }
                                    else if (rightRows.Contains(rowVal))
                                    {
                                        //The row exists in the right sheet but the column only exists in the left meaning the right row will not have a value for this column
                                        cell.SetCellValue("");
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                                    }

                                    colColors[j] = NPOI.HSSF.Util.HSSFColor.Blue.Index;
                                }
                                else if (rightCols.Contains(colVal))
                                {
                                    if (leftRows.Contains(rowVal) && rightRows.Contains(rowVal))
                                    {
                                        cell.SetCellValue(sheetRight.GetRow(dictRightRows[rowVal]).GetCell(dictRightCols[colVal]).StringCellValue);
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Orange.Index;
                                        oranges.Add(rowVal);
                                    }
                                    else if (leftRows.Contains(rowVal))
                                    {
                                        //The row exists in the left sheet but the column only exists in the right meaning the left row will not have a value for this column
                                        cell.SetCellValue("");
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Red.Index;
                                    }
                                    else if (rightRows.Contains(rowVal))
                                    {
                                        cell.SetCellValue(sheetRight.GetRow(dictRightRows[rowVal]).GetCell(dictRightCols[colVal]).StringCellValue);
                                        rowColors[i - 1] = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                                    }

                                    colColors[j] = NPOI.HSSF.Util.HSSFColor.SkyBlue.Index;
                                }
                            }
                        }

                        Console.WriteLine($"MergedWorkbook: Applying colors to Rows&Cols merged worksheet");

                        //Apply colors
                        for (int i = 1; i <= rows; ++i)
                        {
                            for (int j = 0; j < columns; ++j)
                            {
                                int color = rowColors[i - 1];
                                if (color == tempColor)
                                {
                                    color = NPOI.HSSF.Util.HSSFColor.Green.Index;
                                    greens.Add(sheet.GetRow(i).GetCell(0).StringCellValue);
                                }
                                var cell = sheet.GetRow(i).GetCell(j);
                                switch (color)
                                {
                                    case NPOI.HSSF.Util.HSSFColor.Green.Index: cell.CellStyle = greenCellStyle; break;
                                    case NPOI.HSSF.Util.HSSFColor.Red.Index: cell.CellStyle = redCellStyle; break;
                                    case NPOI.HSSF.Util.HSSFColor.Yellow.Index: cell.CellStyle = yellowCellStyle; break;
                                    case NPOI.HSSF.Util.HSSFColor.Blue.Index:
                                        cell.CellStyle = blueCellStyle;
                                        break;

                                    case NPOI.HSSF.Util.HSSFColor.SkyBlue.Index:
                                        cell.CellStyle = cyanCellStyle;
                                        break;

                                    case NPOI.HSSF.Util.HSSFColor.Orange.Index:
                                        cell.CellStyle = orangeCellStyle;
                                        break;

                                    default: break;
                                }

                                //Apply Column Colors. Col colors override row colors
                                if (colColors[j] != tempColor)
                                {
                                    if (colColors[j] == NPOI.HSSF.Util.HSSFColor.Blue.Index)
                                        cell.CellStyle = blueCellStyle;
                                    else if (colColors[j] == NPOI.HSSF.Util.HSSFColor.SkyBlue.Index)
                                        cell.CellStyle = cyanCellStyle;
                                }
                            }
                        }
                        for (int i = 1; i < rows; ++i)
                        {
                            var cell = sheet.GetRow(i).GetCell(0);
                            switch (rowColors[i - 1])
                            {
                                case NPOI.HSSF.Util.HSSFColor.Green.Index: cell.CellStyle = greenCellStyle; break;
                                case NPOI.HSSF.Util.HSSFColor.Red.Index: cell.CellStyle = redCellStyle; break;
                                case NPOI.HSSF.Util.HSSFColor.Yellow.Index: cell.CellStyle = yellowCellStyle; break;
                                case NPOI.HSSF.Util.HSSFColor.Blue.Index:
                                    cell.CellStyle = blueCellStyle;
                                    break;

                                case NPOI.HSSF.Util.HSSFColor.SkyBlue.Index:
                                    cell.CellStyle = cyanCellStyle;
                                    break;

                                case NPOI.HSSF.Util.HSSFColor.Orange.Index:
                                    cell.CellStyle = orangeCellStyle;
                                    break;

                                default: break;
                            }
                        }



                        Console.WriteLine($"MergedWorkbook: Setting metadata in metadata worksheet");
                        row0.CreateCell(2).SetCellValue(reds.Count);
                        row0.CreateCell(3).SetCellValue(yellows.Count);
                        row0.CreateCell(4).SetCellValue(greens.Count);
                        row0.CreateCell(5).SetCellValue(oranges.Count);
                        row0.CreateCell(6).SetCellValue(blues.Count);
                        row0.CreateCell(7).SetCellValue(cyans.Count);

                        rowCount++;
                    }
                    Console.WriteLine($"MergedWorkbook: Finished processing worksheets");
                    rowCount += 2;

                    Console.WriteLine($"MergedWorkbook: Defining missing tables");
                    //Missing Tables
                    var tempRow = sheet0.CreateRow(rowCount++);
                    tempRow.CreateCell(0).SetCellValue("Tables Missing on Left");
                    tempRow.CreateCell(2).SetCellValue("Tables Missing on Right");

                    tempRow = sheet0.CreateRow(rowCount++);
                    tempRow.CreateCell(0).SetCellValue("Sheet name");
                    tempRow.CreateCell(1).SetCellValue("Table name");
                    tempRow.CreateCell(2).SetCellValue("Sheet name");
                    tempRow.CreateCell(3).SetCellValue("Table name");

                    for (int i = 0; i < Math.Max(outLeft.Count, outRight.Count); ++i)
                    {
                        var row = sheet0.CreateRow(i + rowCount + 1);
                        if (i < outLeft.Count)
                        {
                            row.CreateCell(0).SetCellValue(outLeft[i]);
                            row.CreateCell(1).SetCellValue(tableDict[int.Parse(outLeft[i])]);
                            workbookRight.GetSheet(outLeft[i]).CopyTo(workbook, outLeft[i], true, true);
                        }
                        if (i < outRight.Count)
                        {
                            row.CreateCell(2).SetCellValue(outRight[i]);
                            row.CreateCell(3).SetCellValue(tableDict[int.Parse(outRight[i])]);
                            workbookLeft.GetSheet(outRight[i]).CopyTo(workbook, outRight[i], true, true);
                        }
                    }



                }
            }



            for (int i = 0; i < rowCount; ++i)
            {
                sheet0.AutoSizeColumn(i);
            }

            workbook.Write(fs);
            Console.WriteLine($"MergedWorkbook: Finished creating workbook");
        }

    }

    private static Dictionary<int, string> getTableMapping(ref XSSFWorkbook workbookLeft, ref XSSFWorkbook workbookRight)
    {
        var tableDict = new Dictionary<int, string>();
        var metadataWorksheetLeft = workbookLeft.GetSheet(metaDataWorkSheet);
        var metadataWorksheetRight = workbookRight.GetSheet(metaDataWorkSheet);

        Console.WriteLine($"MergedWorkbook: Getting Metadata Table name map from Left");

        for (int i = 3; i < metadataWorksheetLeft.LastRowNum + 1; ++i)
        {
            var tableName = metadataWorksheetLeft.GetRow(i).GetCell(0).StringCellValue;
            var intVal = int.Parse(metadataWorksheetLeft.GetRow(i).GetCell(1).StringCellValue);
            if (tableDict.ContainsKey(intVal) && tableDict[intVal] != tableName)
                throw new Exception($"Table Int Mapping Error: Workbook left has table name {tableName} with id {intVal} however the id is also mapped to {tableDict[intVal]}");
            else
                tableDict[intVal] = tableName;


        }

        Console.WriteLine($"MergedWorkbook: Getting Metadata Table name map from right");
        for (int i = 3; i < metadataWorksheetRight.LastRowNum + 1; ++i)
        {
            var tableName = metadataWorksheetRight.GetRow(i).GetCell(0).StringCellValue;

            var intVal = int.Parse(metadataWorksheetRight.GetRow(i).GetCell(1).StringCellValue);
            if (tableDict.ContainsKey(intVal) && tableDict[intVal] != tableName)
                throw new Exception($"Table Int Mapping Error: Workbook right has table name {tableName} with id {intVal} however the id is also mapped to {tableDict[intVal]}, possibly in the workbook Left");
            else
                tableDict[intVal] = tableName;
        }
        return tableDict;
    }

    private static void createMetaDataWorksheet(ref ISheet sheet0, ref int rowCount)
    {
        Console.WriteLine($"MergedWorkbook: Creating Metadata Worksheet");

        var row00 = sheet0.CreateRow(rowCount++);
        row00.CreateCell(0).SetCellValue("MergeResults");
        row00.CreateCell(1).SetCellValue("Colors");

        sheet0.CreateRow(rowCount++).CreateCell(1).SetCellValue("Red: Rows Only in Left");
        sheet0.CreateRow(rowCount++).CreateCell(1).SetCellValue("Yellow: Rows Only in Right");
        sheet0.CreateRow(rowCount++).CreateCell(1).SetCellValue("Green: Both Row & Column same");
        sheet0.CreateRow(rowCount++).CreateCell(1).SetCellValue("Orange: Same key but differing values");
        sheet0.CreateRow(rowCount++).CreateCell(1).SetCellValue("Blue: Columns Only in Left");
        sheet0.CreateRow(rowCount++).CreateCell(1).SetCellValue("Cyan: Columns Only in Right");

        var row08 = sheet0.CreateRow(rowCount++);
        row08.CreateCell(0).SetCellValue("SheetName");
        row08.CreateCell(1).SetCellValue("TableName");
        row08.CreateCell(2).SetCellValue("No. of Reds");
        row08.CreateCell(3).SetCellValue("No. of Yellows");
        row08.CreateCell(4).SetCellValue("No. of Greens");
        row08.CreateCell(5).SetCellValue("No. of Oranges");
        row08.CreateCell(6).SetCellValue("No. of Blues");
        row08.CreateCell(7).SetCellValue("No. of Cyans");

    }

    private static ICellStyle createCellStyle(ref XSSFWorkbook workbook, short colorIndex)
    {
        var cellStyle = workbook.CreateCellStyle();
        cellStyle.FillForegroundColor = colorIndex;
        cellStyle.FillPattern = FillPattern.SolidForeground;
        cellStyle.BorderLeft = BorderStyle.Thin;
        cellStyle.BorderRight = BorderStyle.Thin;
        cellStyle.BorderTop = BorderStyle.Thin;
        cellStyle.BorderBottom = BorderStyle.Thin;
        cellStyle.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
        cellStyle.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
        cellStyle.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
        cellStyle.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
        return cellStyle;



    }
}
