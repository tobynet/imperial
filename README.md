# Imperial {template engine | subset of mustache | etc}

The `Imperial` is a template engine for `Excel` using `PowerShell`.

todo: Screenshots here

# Requirements

* PowerShell 2.0 or lator
* Windows 7 or lator

# Installation

In PowerShell prompt: 

```
> $modulePath = Join-Path ( $env:psmodulepath -split ';' | select -first 1 ) 'imperial'
> mkdir -Force -Path $modulePath; cd $modulePath
> $client = New-Object System.Net.WebClient
> $client.UseDefaultCredentials = $true
> $client.DownloadFile('https://raw.githubusercontent.com/toooooooby/imperial/master/imperial.psm1', 'imperial.psm1')
```

> :memo: The oepration of the above means **download modules and put in to the directory for module**.

# How to use

1. Import module
    
    ```
    > Import-Module imperial -verbose
    ```

    You dont't have to do above if you have PowerShell 3.0 or lator.

1. Make a template file for Excel.

    * Write `{{variable_name}}` in Excel files

    todo: Write how to make template with some screenshots.

1. Use it.
    
    Simple:

    ```
    > $variables = @{foo=1234; bar="buzz"}
    > Render-Excel template.xslx out.xslx $variables
    ```

    More details: 

    ```
    > Render-Excel `
        -inputFilename (ls 'tempalte.xslx').fullname `
        -outputFilename ([Io.Path]::GetFullPath("out.xslx")) `
        -variables (Get-HashTableFromCSV (Import-Csv "my_variables.csv"))
    ```

    todo: Screenshots for result here


# Todos

* Write documents
    * Animation GIFs
* **DONE: Modulize**
* Add UnitTests

# Help us!

* You know how to test at PowerShell?
* You know some CI server for testing PowerShell?

