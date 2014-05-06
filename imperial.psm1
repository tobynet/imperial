<#
e.g.
    
    > Render-Template -source "{{foo}} bar" -variables @{foo="REPLACED"}
    REPLACED bar
    
#>
function Render-Template([string]$source, [Hashtable]$variables) {
    [regex]::Replace( $source, "{{([\w_]+)}}", {
            param($matched)
            $variables[$matched.Groups[1].Value]
        })
}

<#
e.g.
    > Get-Content foo.csv
    foo,123
    bar,456

    > Get-HashTableFromCSV (Import-Csv "foo.csv")
    Name           Value
    ----           -----
    foo            123
    bar            456
#>
function Get-HashTableFromCSV([Object[]]$csv) {
    $variables = @{}
    $csv | %{ $variables[$_.key] = $_.value }
    $variables
}


<#
e.g.1

    > Render-Excel -inputFilename "template.xsl" -outputFilename "out.xsl" -variables @{foo="bar",buzz=1234}

e.g.2

    > $hash = Get-HashTableFromCSV (Import-Csv "foo.csv")
    > Render-Excel -inputFilename "template.xsl" -outputFilename "out.xsl" -variables $hash
    
#>
function Render-Excel([string]$inputFilename, [string]$outputFilename, [HashTable]$variables) {
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
                $buffer = Render-Template $buffer $variables
            } else {
                1..$range.Rows.Count | %{ $y = $_
                    1..$range.Columns.Count | %{ $x = $_
                        #echo ("  [{0},{1}] : {2}" -f $y, $x, $buffer[$y, $x])
                        $buffer[$y, $x] = Render-Template $buffer[$y, $x] $variables
                    }
                }
            }
            
            
            $range.Value2 = $buffer
        }
        echo ( "Save as $outputFilename")
        $books.SaveAs($outputFilename)
    } finally {
        $excel.Quit()
    }
}

<#
Render-Excel `
    -inputFilename (ls $inputFile).fullname `
    -outputFilename ([Io.Path]::GetFullPath($outputFile)) `
    -variables (Get-HashTableFromCSV (Import-Csv $csvFile))
echo "DONE"
#>

Export-ModuleMember -Function Render-Excel,Render-Template

