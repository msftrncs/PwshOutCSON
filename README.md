# PwshOutCSON - ConvertTo-CSON

Convert a PowerShell object to a CSON string notation.

Intended to aid in the conversion of tmLanguage grammar definitions to CSON format.

Example based on JSON tmLanguage file.

```PowerShell
Get-Content "powershell.tmlanguage.json" | ConvertFrom-Json |
    ConvertTo-Cson -Indent "`t" -Depth 100 |
    Set-Content 'grammars\PowerShell.cson' -Encoding 'UTF8'
```

```PowerShell
$grammar_json = Get-Content "powershell.tmlanguage.json" | ConvertFrom-Json
ConvertTo-Cson $grammar_json -Indent "`t" -Depth 100 |
    Set-Content 'grammars\PowerShell.cson' -Encoding 'UTF8'
```

```PowerShell
$grammar_cson_doc = Get-Content "powershell.tmlanguage.json" | ConvertFrom-Json |
    ConvertTo-Cson -Indent "`t" -Depth 100
```

### Parameters

#### InputObject

The PowerShell object which possesses the items to be output in CSON notation.  This can be any object that can be represented as a PSCustomObject.  This parameter may be received from the pipeline.

#### Indent

Specifies the indentation to use when generating the CSON output.  The default is ```"`t"``` (tab), but other usual options are `''` (none), `'    '` (4 spaces), but otherwise any string is accepted, and no escaping is performed.

#### Depth

Specifies the depth of recursion permitted for the input object.  See the help for `ConvertTo-JSON` for details.

#### EnumsAsStrings

A switch that specifies an alternate serialization option that converts all enumerations to their string representations.

### Output

The output of this function is a string (`[string]`).  Use Out-File or Set-Content and be sure to assign the correct encoding.  

Note:
- The CSON serialization is only as complete as was required for tmLanguage grammar files that were tested.
- Minimal escaping has been programmed so far.
