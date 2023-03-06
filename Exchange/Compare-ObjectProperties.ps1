param (
    [Parameter(Mandatory = $true)][string]$cmdlet,
    [Parameter(Mandatory = $true)][string]$object1,
    [Parameter(Mandatory = $true)][string]$object2,
    [string]$switches
)

Function checkCmdlet{
    Param(
        [Parameter(Mandatory = $true)][string]$cmdletToTest
    )

    try{
        $testCommand = Get-Command $cmdletToTest -ErrorAction Stop | Out-Null
    } catch {
        Write-Host -foregroundcolor red "Cmdlet '$cmdletToTest' doesn't appear to exist."
        Write-Host -foregroundcolor Red $error[0].Exception.Message
        exit
    }
}

Function InvokePowerShellCmdlet
{
    Param(
        [Parameter(Mandatory = $true)]
        [String]$CmdLetToRun
    )
    try
    {
        return Invoke-Expression $CmdLetToRun -ErrorAction Stop
    }
    catch
    {
        write-Host -ForegroundColor Red "Error executing the following cmdlet:"
        write-Host -foregroundcolor Red $CmdLetToRun
        Write-Host -foregroundcolor Red $error[0].Exception.Message
    }
}

checkCmdlet $cmdlet

$obj1Props = InvokePowerShellCmdlet "$cmdlet $object1 $switches"
$obj2Props = InvokePowerShellCmdlet "$cmdlet $object2 $switches"

if((($obj1Props | Measure).count -eq 1) -and (($obj2Props | Measure).count -eq 1)){

    $ds = New-Object System.Data.DataSet
    $mismatchItems = New-Object System.Data.DataTable
    $mismatchItems.TableName = "Mismatches"

    $mismatchItems.Columns.Add("Property") | Out-Null
    $mismatchItems.Columns.Add("Object 1") | Out-Null
    $mismatchItems.Columns.Add("Object 2") | Out-Null

    $objProps = $obj1Props | Get-Member -MemberType Property,NoteProperty
    $rows
    foreach($prop in $objProps){
        $propName = $prop.name
        if($obj1Props.$propName -ne $obj2Props.$propName){
            $mismatchRow = $mismatchItems.NewRow()
            $mismatchRow.Property = $propName
            $mismatchRow."Object 1" = $obj1props.$propName
            $mismatchRow."Object 2" = $obj2props.$propName
            $mismatchItems.Rows.Add($mismatchRow)
        }
    }
    $ds.Tables.Add($mismatchItems)
    $ds.Tables["Mismatches"] | Format-List

} elseif ((($obj1Props | Measure).count -gt 1) -or (($obj2Props | Measure).count -gt 1)){
    write-Host -ForegroundColor Red "Multiple objects matched. Be more specific."
} else {
    write-Host -ForegroundColor Red "Error getting object properties:"
    Write-Host -foregroundcolor Red $error[0].Exception.Message
}

