<#
.SYNOPSIS
    Convert a PowerShell object to a CSON (Coffee Script) representation in a string.
.DESCRIPTION
    Converts a PowerShell object to a CSON (Coffee Script) notation as a string.
.PARAMETER InputObject
    The input PowerShell object to be represented in a CSON notation.  This parameter may be received from the pipeline.
.PARAMETER Indent
    Specifies a string value to be used for each level of the indention within the CSON document.
.PARAMETER Depth
    Specifies the maximum depth of recursion permitted for the input object.
.PARAMETER EnumsAsStrings
    A switch that specifies an alternate serialization option that converts all enumerations to their string representations.
.EXAMPLE
    $grammar_json | ConvertTo-Cson -Indent `t -Depth 100 | Set-Content out\PowerShell.cson -Encoding UTF8
.INPUTS
    [object] - any PowerShell object.
.OUTPUTS
    [string] - the input object returned in a CSON notation.
.NOTES
    Script / Function / Class assembled by Carl Morris, Morris Softronics, Hooper, NE, USA
    Initial release - Mar 3, 2019
.LINK
    https://github.com/msftrncs/PwshOutCSON/
.FUNCTIONALITY
    data format conversion
#>
function ConvertTo-Cson {
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [AllowNull()]
        [object] $InputObject,

        [PSDefaultValue(Help = 'Tab')]
        [string] $Indent = "`t",

        [ValidateRange(1, 100)]
        [int32] $Depth = 2,

        [switch] $EnumsAsStrings
    )
    # write out a CSON document from the object supplied
    # $InputObject is an object, who's properties will be output as CSON.  Hash tables are supported.
    # $Indent is a string representing the indentation to use.
    #   Typically use "`t" or "  ".

    # define a match evaluator for escaping characters
    $escape_replacer = {
        if ($_.Groups[1].Success) {
            # group 1, control characters
            switch ($_.Value[0]) {
                <# appearing in order of expected frequency, from most frequent to least frequent #>
                ([char]10) { '\n'; continue } # new line
                ([char]9) { '\t'; continue }  # tab
                ([char]13) { '\r'; continue } # carriage return
                ([char]12) { '\f'; continue } # new form
                ([char]8) { '\b'; continue }  # bell
                default { '\u{0:X4}' -f [int16]$_ }   # unicode escape all others
            }
        } elseif ($_.Groups[2].Success) {
            # group 2, items that need `\` escape
            "\$($_.Value)"
        }
    }

    filter writeStringValue {
        # write an escaped CSON string property value
        # the purpose of making this a function, is a single place to change the escaping function used
        # TODO: escape more characters!
        """$($_ -replace '([\x00-\x1F\x85\u2028\u2029])|([\\"]|#\{)', $escape_replacer)"""
    }

    function writeObject ($item) {

        function writeProperty ([string] $name, $value) {
            # write a property name and its value, which may require recursing back to writeObject
            "$indention$(
                        # if a property name is not all simple characters or starts with numeric digit, it must be quoted and escaped
                        if (-not $name -or $name -match '[^\p{L}\d_]|^\d') {
                            # property name requires escaping
                            $name | writeStringValue
                        }
                        else {
                            $name
                        }
                    ):$(
                        if (($level -gt $Depth) -or ($null -eq $value) -or ($value -is [valuetype]) -or ($value -is [string])) {
                            " $(, $value | writeValue)" # comma forces $value to be treated as a whole object instead of being enumerated as subitems
                        }
                        elseif ($value -is [Collections.IList]) {
                            " [$( # add array start token if value is an array
                                    if ($value.Count -eq 0) {
                                        ']' # add array end token if value is an empty array
                                    }
                                )"
                        }
                    )"
            if ($level -le $Depth) {
                # if exceeded Depth, value already written above
                if (($value -is [Collections.IList]) -and ($value.Count -ne 0)) {
                    # handle nested non-empty arrays specially due to already emitted array start token
                    $level++ # level increases for arrays or objects
                    $value | writeArray # write the nested array
                    "$indention]" # nested array end token
                } elseif ($value -and ($value -isnot [valuetype]) -and ($value -isnot [string])) {
                    $indention = "$indention$Indent"
                    writeObject $value # recurse the element to writeObject
                }
            }
        }

        filter writeValue {
            # write a object property or array element simple value
            if ($null -eq $_) {
                'null'
            } elseif (($_ -is [char]) -or ($EnumsAsStrings -and ($_ -is [enum])) -or ($_ -isnot [valuetype])) {
                # handle strings or characters, or objects exceeding the max depth
                "$_" | writeStringValue
            } elseif ($_ -is [boolean]) {
                # handle boolean type
                if ($_) {
                    'true'
                } else {
                    'false'
                }
            } elseif ($_ -is [datetime]) {
                # specifically format date/time to ISO 8601
                $_.ToString('o') | writeStringValue
            } elseif ($_ -isnot [enum]) {
                # assuming a [valuetype] that doesn't need special treatment
                $_
            } else {
                # specifically out the enum value
                $_.value__
            }
        }

        filter writeArray {
            begin {
                $indentionArray = "$indention$Indent"
            }
            process {
                # if depth not exceeded, check for a nested object
                if (($level -le $Depth) -and $_ -and ($_ -isnot [valuetype]) -and ($_ -isnot [string]) -and ($_ -isnot [Collections.IList])) {
                    # an object is nested within the array element
                    if ($(if ($_ -is [Collections.IDictionary]) { $_.get_Keys().Count } else { @($_.psobject.get_Properties()).Count } ) -gt 0) {
                        "$indentionArray{" # object start token
                        $indention = "$indentionArray$Indent"
                        writeObject $_ # recurse the object to writeObject
                        "$indentionArray}" # object end token
                    } else {
                        "$indentionArray{}" # empty object
                    }
                } else {
                    # for all other cases recurse the value back to writeObject
                    $indention = $indentionArray
                    writeObject $_
                }
            }
        }

        if (($level -gt $Depth) -or ($null -eq $item) -or ($item -is [valuetype]) -or ($item -is [string])) {
            "$indention$(, $item | writeValue)" # comma forces $item to be treated as a whole object instead of being enumerated as subitems
        } else {
            $level++ # level increases for arrays or objects
            if ($item -is [Collections.IList]) {
                if ($item.Count -ne 0) {
                    # handle non-empty arrays, iterate through the items in the array
                    "$indention["
                    $item | writeArray
                    "$indention]"
                } else {
                    "$indention[]" # indicate empty array
                }
            } elseif ($(if($item -is [Collections.IDictionary]) { $item.get_Keys().Count } else { @($item.psobject.get_Properties()).Count } ) -gt 0) {
                if ($item -is [Collections.IDictionary]) {
                    # process what we assume is a hashtable object
                    foreach ($key in $item.get_Keys()) {
                        # handle objects by recursing with writeProperty
                        writeProperty $key $item[$key]
                    }
                } else {
                    # iterate through the objects properties
                    foreach ($property in $item.psobject.Properties) {
                        # handle objects by recursing with writeProperty
                        writeProperty $property.Name $property.Value
                    }
                }
            } else {
                "$indention{}" # indicate empty object
            }
        }
    }

    # start writing the input object starting with no indent at level 0
    [string] $indention = ''
    [int32] $level = 0
    (writeObject $(
                # we need to determine where our input is coming from, pipeline or parameter argument.
                if ($input -is [array] -and $input.Length -ne 0) {
                    $input # input from pipeline
                } else {
                    $InputObject # input from parameter argument
                }
            )
        ) -join "$(if (-not $IsCoreCLR -or $IsWindows) { "`r" })`n"
}
