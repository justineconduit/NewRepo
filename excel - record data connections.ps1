###################################################################################################
#
#	Name:			excel - record data connections.ps1
#	Description:	This will capture data connection of xerox reports and record it to the database
#					Xerox reports are excel files that uses SQL Command in the data connection
#	Author:			Anthony O. Tabernero (Thonz)
#	Created:		16.Mar.2020
#
#	Note:			It is best to run this script on host where the sharepoint files are downloaded
#					Host = TX0CMSRDB01; DB = _ETL_IS_Mgt; Local Copy = C:\kgb_\xerox uk\insights
#
###################################################################################################
Clear-Host

#$var_rootDIR = "C:\kgb_\xerox uk\insights\xerox\ReportsLibrary\CSC - Xerox MI\ECH Data\*.xlsx"
$var_rootDIR = "C:\kgb_\xerox uk\insights\xerox\ReportsLibrary\*.xlsx" #(If run from diff machine = "\\Tx0cmsrdb01\c$\kgb_\xerox uk\insights\xerox\ReportsLibrary\*.xlsx)"

$dataSource = "TX0CMSRDB01"
$database = "_ETL_IS_Mgt"
$connectionString = "Server=$dataSource;Database=$database;Integrated Security=True;"

Function Get-DataConnection
{
    Param([String]$input_Excel)
    $result = @()

    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.visible =$False
    $objExcel.DisplayAlerts = $False

    $xlsWorkBook = $objExcel.workbooks.Open($input_Excel)
    $connections = $xlsWorkBook.Connections

    foreach ($conn in $connections)  
    {     
        if ($conn -ne $null)  
        {  
            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'conn' -Value $conn.Name
            $item | Add-Member -type NoteProperty -Name 'sqlCmd' -Value $conn.OLEDBConnection.CommandText

            $result += $item

        }  
    } 
    $xlsWorkBook.Close($False)
    $objExcel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
    Remove-Variable xlsWorkBook,objExcel | Out-Null

    return($result)
}


Write-Host "Scanning Xerox (UK) reports' data connections"
Write-Host "Directory = $var_rootDIR"
Write-Host "Please wait ..."
Write-Host ""
Start-Sleep -s 1.5

$files = Get-ChildItem -Path $var_rootDIR –Recurse -File #| Where-Object{$_.Name -contains ".xlsx"}

foreach($file in $files)
{
    $file_rpt = $file.name.Substring(0, $file.name.IndexOf(".xlsx"))
    $file_dir = $file.Directory.Tostring()
    $file_src = $file.FullName

    $len_filedir = $file_dir.Length
    $indx_xerox = $file_dir.IndexOf("xerox\")        
    $itemdir = $file_dir.Substring($indx_xerox, $len_filedir - $indx_xerox)
    
    Write-Host "     Capturing data connection definition for ""$file_rpt"" ..."

    $dataConn = Get-DataConnection -input_Excel $file.FullName
  
    foreach($conn in $dataConn)
    {
        $conn_name = $conn.conn
        $conn_sqlcmd = $conn.sqlCmd


        $query = "INSERT INTO xeroxUK.ReportReference_DataConn
                    (
                        ReportName,
                        DataConn,
                        SQLCommand,
                        ItemDirectory,
                        FileSource
                    )
                    VALUES
                    (
                       @report,
                       @conn,
                       @sqlcmd,
                       @itemdir,
                       @filesrc
                    )"
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

            $command = $connection.CreateCommand()
            $command.CommandText = $query
        
            $command.Parameters.Add("@report", $file_rpt)  | Out-Null
            $command.Parameters.Add("@conn", $conn_name)  | Out-Null
            $command.Parameters.Add("@sqlcmd", $conn_sqlcmd)  | Out-Null
            $command.Parameters.Add("@itemdir", $itemdir)  | Out-Null
            $command.Parameters.Add("@filesrc", $file_src)  | Out-Null

            $command.ExecuteNonQuery() | Out-Null

        $connection.Close()
    }
    
   
}

Write-Host ""
Write-Host "Capture Complete!"
Write-Host "Data are stored at ""xeroxUK.ReportReference_DataConn"""
