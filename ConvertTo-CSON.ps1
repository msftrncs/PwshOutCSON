# write out an object in CSON notation.

# define a function to create a CSON document, trying to keep it as generic as possible
function ConvertTo-Cson {
    param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [AllowNull()]
        [object]$InputObject,

        [string]$Indent = "`t",

        [ValidateRange(1, 100)]
        [int32]$Depth = 2,

        [switch]$EnumsAsStrings
    )
    # write out a CSON document from the object supplied
    # $InputObject is an object, who's properties will be output as CSON.  Hash tables are supported.
    # $Indent is a string representing the indentation to use.
    #   Typically use "`t" or "  ".

    # define a REGEX match for simple property names
    $match_isNotSimpPropName = [regex]::new('[^\p{L}\d_]|^\d')
    # define a REGEX match for characters or patterns that need escaped
    $escape_matcher = [regex]::new('([\x00-\x1F\x85\u2028\u2029])|([\\"]|#\{)')
    # define a match evaluator for escaping characters
    $escape_replacer = {
        switch ($args[0]) {
            {$_.Groups[1].Success} {
                # group 1, control characters
                switch ($_.Value[0]) {
                    ([char]8) {'\b'}
                    ([char]9) {'\t'}
                    ([char]10) {'\n'}
                    ([char]12) {'\f'}
                    ([char]13) {'\r'}
                    default {'\u{0:X4}' -f [int16]$_}
                }
            }
            {$_.Groups[2].Success} {
                # group 2, items that need `\` escape
                "\$($_.Value)"
            }
        }
    }

    function writeStringValue ([string]$value) {
        # write an escaped CSON string property value
        # the purpose of making this a function, is a single place to change the escaping function used
        # TODO: escape more characters!
        """$($escape_matcher.replace($value, $escape_replacer))"""
    }

    function writePropertyName ([string]$value, [bool]$isArray) {
        # write an property name, processed as required for CSON
        # the purpose of making this a function, is a single place to change the escaping function used
        $(
            if ($match_isNotSimpPropName.IsMatch($value)) {
                # property name requires escaping
                "$(writeStringValue $value)"
            }
            else {
                $value
            }
        ) + ':' + $(
            if ($isArray) {
                " ["
            }
        )
    }
    function writeproperty ([string]$name, $item, [string]$indention, [int32]$level) {
        # writing the property may require recursively breaking down the objects based on their type
        # name of the property is optional, but that is only intended for the first property object
        function writevalue ($item, [string]$indention) {
            # write a property value

            if (($item -is [string]) -or ($item -is [char])) {
                # handle strings or characters
                "$indention$(writeStringValue $item)"
            }
            else {
                "$indention$(
                    if ($item -is [boolean]) {
                        # handle boolean type
                        if ($item) {
                            "true"
                        }
                        else {
                            "false"
                        }
                    } elseif ($item -isnot [enum]) {
                        "$item"
                    } elseif ($EnumsAsStrings) {
                        writeStringValue ($item.ToString())
                    } else {
                        "$($item.value__)"
                    }
                )"
            }
        }

        if ($level -le $Depth) {
            # write out key name, if one was supplied from the parent object
            if ($name) {
                "$indention$(writePropertyName $name ($item -is [array]))" + $(
                    if ($item -is [ValueType] -or $item -is [string]) {" $(writevalue $item """")"}
                )
            }
            else {
                if ($item -is [array]) {
                    "$indention[" # add array start token if property is an array
                }
                elseif ($item -is [valuetype] -or $item -is [string]) {
                    writevalue $item "$indention"
                }
            }

            if ($item -is [array]) {
                # handle arrays, iterate through the items in the array
                foreach ($subitem in $item) {
                    if ($subitem -is [valuetype] -or $subitem -is [string]) {
                        writevalue $subitem "$indention$Indent"
                    }
                    else {
                        "$indention$indent{"
                        writeproperty $null $subitem "$indention$Indent" ($level + 1)
                        "$indention$Indent}"
                    }
                }
                "$indention]"
            }
            elseif ($item -isnot [valuetype] -and $item -isnot [string]) {
                # handle objects by recursing with writeproperty
                # iterate through the items (force to a PSCustomObject for consistency)
                foreach ($property in ([PSCustomObject]$item).psobject.Properties) {
                    writeproperty $property.Name $property.Value $(if ($level -ge 0) {"$indention$Indent"} else {$indention}) ($level + 1)
                }
            }
        }
        else {
            # exceeded maximum depth, convert object to string
            "$indention$(writeStringValue $item)"
        }
    }

    # start writing the property list, the property list should be an object, has no name, and starts at base level
    (writeproperty $null $InputObject "" (-1)) -join "`r`n"
}
