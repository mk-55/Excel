# 
function Get-WorkSheetNames {
    [CmdletBinding()] #Write As a ScriptCmndlet
    Param (
        [parameter(mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Path,

        # todo need more better naming
        [parameter(mandatory = $false)]
        [boolean]$needsOutputFileName = $false
    )

    begin {
        # open Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

    }

    process {
        try {
            # Argument Validation
            if (-not(Test-Path $Path -Include "*.xlsx", "*.xls", "*.xlsm" )) {
                Write-Debug "File not Found or File not ExcelBook"
                throw  New-Object "System.ArgumentException" @("File not Found or not ExcelBook", $Path) -ErrorAction Stop
            }

            # Get absolute path (for Excel.Application.Workbooks.Open())
            $ExcelPath = Convert-Path $Path
            
            # Get sheet name
            $book = $excel.Workbooks.Open($ExcelPath)
            if ($needsOutputFileName) {
                $fileName = Split-Path $ExcelPath -Leaf
                $book.WorkSheets | ForEach-Object {Write-Host $filename " : " $_.Name }
            } else {
                $book.WorkSheets | ForEach-Object { $_.Name }
            }
            
        
        } catch [System.ArgumentException] {
            Write-Error $_

        } catch {
            # in order to close Excel in End Block
            Write-Errot $_
        
        } finally {
            # close book
            if ($book) {
                [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
            }
        }
    }

    end {
        # close Excel
        Write-Debug "Close Excel"
        [void]$excel.Quit()
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        
    }
}

Export-ModuleMember -Function Get-WorkSheetNames