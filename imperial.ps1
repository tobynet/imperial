# Usage: imperial.ps1 input.xslx output.xslx variables.csv

Param (
    [parameter(Mandatory=$true)][string]$inputFile,
    [parameter(Mandatory=$true)][string]$outputFile,
    [string]$csvFile
)

function renderTemplate($source, $variables) {
    [regex]::Replace( $source, "{{([\w_]+)}}", {
            param($matched)
            $variables[$matched.Groups[1].Value]
        })
}

function getHashTableFromCSVFile($csvFile) {
    $variables = @{}
    Import-Csv $csvFile | %{ $variables[$_.key] = $_.value }
    $variables
}

function renderExcel($inputFilename, $outputFilename, $variables) {
    echo ( "Rendering an excel file from {0} to {1}" -f $inputFilename , $outputFilename )
    
    # Making an Excel object
    # $excel.Visible = $false           # Invisible!!
    # $excel.DisplayAlerts = $false     # Ignore dialogs for erros
    $excel = New-Object -com Excel.Application -Property @{visible=$false; DisplayAlerts=$false}
    try {
        # Overwrite when the file is saved
        $excel.AlertBeforeOverwriting = $true
        
        $books = $excel.Workbooks.Open($inputFilename)
        
        # todo: Maybe this block should rather use `foreach($x in $xs)` instead of `foreach-object`...
        $books.Sheets | %{
            echo ( "Replacing sheet... : {0}" -f $_.name )
            $sheet = $_
            $range = $sheet.usedRange

            # Value2 is a fast member which can manupulate excel's cells!!
            # It's TOO slow if `sheet.cells.item(y,x)` is used
            $buffer = $range.Value2
            
            if ($range.Rows.Count -eq 1) {
                # It is a NOT 2-dimensions array if count of Rows equls 1 !?!?!?????
                $buffer = renderTemplate $buffer $variables
            } else {
                1..$range.Rows.Count | %{ $y = $_
                    1..$range.Columns.Count | %{ $x = $_
                        #echo ("  [{0},{1}] : {2}" -f $y, $x, $buffer[$y, $x])
                        $buffer[$y, $x] = renderTemplate $buffer[$y, $x] $variables
                    }
                }
            }
            
            
            $range.Value2 = $buffer
        }
        echo ( "Save as {0}" -f $outputFilename )
        $books.SaveAs($outputFilename)
    } finally {
        $excel.Quit()
    }
}

renderExcel $(ls $inputFile).fullname $([Io.Path]::GetFullPath($outputFile)) $(getHashTableFromCSVFile $csvFile)

echo "DONE"
