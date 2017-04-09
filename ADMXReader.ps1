<#
.SYNOPSIS
  Citrix ADMX Policy information converter.
.DESCRIPTION
  The Citrix ADMX Policy information converter converts the ADMX XML format to a more readable CSV file with policy and registry information.
.PARAMETER ADMXFolder
    The Folder containing the ADMX files that will be processed by this script. Please include a language subfolder for the corresponding ADML files.
.PARAMETER OutputCSVFile
    The name (and extension) of the output CSV file created by the script.
.PARAMETER language
    The language (and subfolder name) of the corresponding ADML files to be used by the script. Only en-US has been used during script testing and is included in the selectable languages.
.OUTPUTS
  Policy information CSV file stored in the root of the provided policy directory.
.NOTES
  Version:        1.3
  Author:         Esther Barthel, MSc
  Creation Date:  2017-02-10
  Purpose/Change: Initial script development
  Update Date:    2017-03-19
  Purpose/Change: Adjustments after first Windows ADMX files run
  Update Date:    2017-03-23
  Purpose/Change: Adjustments to proces policies with values correctly in their own rows
  Update Date:    2017-03-31
  Purpose/Change: Added more possible values data

  Copyright (c) cognition IT. All rights reserved.

  Based upon this reference material: https://technet.microsoft.com/en-us/library/cc731761(v=ws.10).aspx
                                      https://technet.microsoft.com/en-us/library/cc771659(v=ws.10).aspx (ADMX syntax)
                                      https://technet.microsoft.com/en-us/library/cc753471(v=ws.10).aspx (Group Policy ADMX Syntax Reference Guide)
                                      https://msdn.microsoft.com/en-us/library/dn606024(v=vs.85).aspx (more recent syntax) 
.EXAMPLE
  ADMXReader_v1_3.ps1 -ADMXFolder $($env:windir)\policyDefinitions -OutputCSVFile "ADMXOutput.csv" -language en-US
#>
[CmdletBinding()]
# Declaring script parameters
Param(
    [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()][string]$ADMXFolder="C:\temp\ADMXReader\ADMXFiles",
    [Parameter(Mandatory=$true)] [ValidateNotNullOrEmpty()][string]$OutputCSVFile="C:\temp\ADMXReader\Script_Output_Test.csv",
    [Parameter(Mandatory=$true)] [ValidateSet("en-US")] [string]$language="en-US"
)
#requires -version 3

Clear-Host
#region---------------------------------------------------------[Initialisations]--------------------------------------------------------
    $rowCount = 0
    $policyCount = 0
    $elementCount = 0
    $dropdownListValues = ""
    $itemCount = 0
#endregion

#region----------------------------------------------------------[Declarations]----------------------------------------------------------
    # Creating the DataTable object
    $table= New-Object System.Data.DataTable

    # Setting the table headers
    [void]$table.Columns.Add("ADMX")
    [void]$table.Columns.Add("Parent Category")
    [void]$table.Columns.Add("Name")
    [void]$table.Columns.Add("Display Name")
    [void]$table.Columns.Add("Class")
    [void]$table.Columns.Add("Explaining Text")
    [void]$table.Columns.Add("Supported On")
    [void]$table.Columns.Add("Type")
    [void]$table.Columns.Add("Label")
    [void]$table.Columns.Add("Registry Key")
    [void]$table.Columns.Add("Value Name")
    [void]$table.Columns.Add("Possible Values")

    If ($PSBoundParameters['Debug']) 
    {
        # Changing Debugging from default 'Inquire' setting (with prompted actions per Write-Debug line) to continu to generate Debug messages without the prompts.
        $DebugPreference = [System.Management.Automation.ActionPreference]::Continue
    }
#endregion

#region-----------------------------------------------------------[Execution]------------------------------------------------------------

If (!(Test-Path -Path $ADMXFolder))
{
    Throw "Policy Directory $ADMXFolder NOT Found. Script Execution STOPPED."
}

# Retrieving al the ADMX files to be processed
$admxFiles = Get-ChildItem $ADMXFolder -filter *.admx

Write-Output ($admxFiles.Count.ToString() + " ADMX files found in """ + $ADMXFolder + """")

# Checking for the Windows and Citrix supportedOn vendor definition files
If (Test-Path("$ADMXFolder\$language\Windows.adml"))
{
    [xml]$supportedOnWindowsTableFile = Get-Content "$ADMXFolder\$language\Windows.adml"
}
If (Test-Path("$ADMXFolder\$language\CitrixBase.adml"))
{
    [xml]$supportedOnCitrixTableFile = Get-Content "$ADMXFolder\$language\CitrixBase.adml"
}

ForEach ($file in $admxFiles)
{
    #Proces each file in the directory
    Write-Output ("*** Processing file " + $file.Name)
 
    [xml]$data=Get-Content "$ADMXFolder\$($file.Name)"
    [xml]$lang=Get-Content "$ADMXFolder\$language\$($file.Name.Replace(".admx",".adml"))"

    # Retrieve all information groups from the specific ADMX file (for easy searches)
    $stringTableChilds = $lang.policyDefinitionResources.resources.stringTable.ChildNodes
    $presentationTableChilds = $lang.policyDefinitionResources.resources.presentationTable.ChildNodes
    $supportedOnDefChilds = $data.policyDefinitions.supportedOn.definitions.ChildNodes
    $categoryChilds = $data.policyDefinitions.categories.ChildNodes

    #Initializing the policy Counter
    $policyCount = 0

    # Processing each policy (ChildNode) of the policies Node
    $data.PolicyDefinitions.policies.ChildNodes | ForEach-Object {
        # Get current policy node
        $policy = $_
        If ($policy -ne $null)
        {
            # Processing policy nodes with a name, other than #comment (making sure the comment tags in the file are ignored!)
            If ($policy.Name -ne "#comment")
            {
                $policyCount = $policyCount + 1
                Write-Output "* Processing policy $($policy.Name)"

                # Retrieving policy information from node
                $polDisplayName = ($stringTableChilds | Where-Object { $_.id -eq $policy.displayName.Substring(9).TrimEnd(')') }).InnerText          # Getting the displayName from the StringTable 
                Write-Debug "displayName for $($policy.Name) is ""$polDisplayName"""
                $explainText = ($stringTableChilds | Where-Object { $_.id -eq $policy.explainText.Substring(9).TrimEnd(')') }).InnerText             # Getting the explainText from the StringTable 
                Write-Debug "explainText for $($policy.Name) is ""$explainText"""
                $regkey = $policy.key 
                Write-Debug "registry key for $($policy.Name) is ""$regkey"""

                #region retrieving supportedOn Childnode information              
                If ($policy.SupportedOn.ref.Contains(":"))
                {        
                    $supportedOnVendor=$policy.SupportedOn.ref.Split(":")[0]

                    Switch ($supportedOnVendor.ToLower())
                    {
                        "windows"
                        {
                            # Using the Windows supportedOn information from the Windows.ADMX file
                            If ($supportedOnWindowsTableFile -ne $nulll)
                            {
                                $supportedOnTableChilds = $supportedOnWindowsTableFile.policyDefinitionResources.resources.stringTable.ChildNodes
                            }
                        }
                        "citrix"
                        {
                            # Using the Citrix supportedOn information from the Citrix.ADMX file
                            If ($supportedOnCitrixTableFile -ne $null)
                            {
                                $supportedOnTableChilds = $supportedOnCitrixTableFile.policyDefinitionResources.resources.stringTable.ChildNodes
                            }
                        }
                        default
                        {
                            # Use the specific supportedOn information from the current ADMX file
                            $supportedOnTableChilds = $stringTableChilds
                        }
                    }
                    $supportedOnID=$policy.SupportedOn.ref.Split(":")[1]
                    $supportedOn=($supportedOnTableChilds | Where-Object { $_.id -eq $supportedOnID }).InnerText
                    If ([string]::IsNullOrEmpty($supportedOn))
                    {
                        $supportedOn=($stringTableChilds | Where-Object { $_.id -eq $supportedOnID }).InnerText
                    }
                }
                Else
                # no ':' in supportedOn information, find name in right supportedOn identity
                {
                    $supportedOnID = ($supportedOnDefChilds | Where-Object { $_.Name -eq $policy.supportedOn.ref }).displayName
                    If ($supportedOnID -ne $null)
                    {
                        # supportedOn displayname found, find description text in stringTable ChildNodes
                        $supportedOn = ($stringTableChilds | Where-Object { $_.id -eq $supportedOnID.Substring(9).TrimEnd(')') }).InnerText
                    }
                    Else
                    {
                        # supported on information not found
                        $supportedOn = "*unknown*"
                    }
                }
                Write-Debug "SupportedOn for $($policy.Name) is ""$supportedOn"""
                #endregion retrieving supportedOn information 

                #region retrieving parentCategory Childnode information
                If ($policy.parentCategory.ref.Contains(":"))
                {
                    $parentCategoryID=$policy.parentCategory.ref.Split(":")[1]
                    $parentCategory=($stringTableChilds | Where-Object { $_.id -eq $parentCategoryID }).InnerText
    
                } 
                Else
                # no ':' in categoryParent information, find name in right Category identity
                {
                     $parentCategoryID =  ($categoryChilds | Where-Object { $_.Name -eq $policy.parentCategory.ref }).displayName
                     If ($parentCategoryID -ne $null)
                     {
                         # parentCategory displayname found, find description text in stringTable ChildNodes
                         $parentCategory =  ($stringTableChilds | Where-Object { $_.id -eq $parentCategoryID.Substring(9).TrimEnd(')') }).InnerText
                     }
                     Else
                     {
                        # Check the ADML file for Category Label
                         $parentCategory =  ($stringTableChilds | Where-Object { $_.id -eq $policy.parentCategory.ref }).'#text'
                     }
                }
                Write-Debug "parentCategory for $($policy.Name) is ""$parentCategory"""
                #endregion retrieving parentCategory information

# Include a Row with policy information (and 1-Check for existing policy valueName ; 2-start a new row after this

                #region retrieve policy attribute information (valueName, etc) and log a row for policy information (linked to the enabledValue and disabledValue subelements)
                If ($policy.valueName -ne $null)
                {
                    Write-Host ("Policy value " + $policy.valueName + " found") -ForegroundColor Cyan

                    # Note: BIG ASSUMPTION: policy valueName is always processed as an Enable/Disable boolean registry key value
                    $polValueName = $policy.valueName

                    $polPossibleValues = "Policy value:"
                    If ($policy.enabledValue -ne $null)
                    {
                        $polEnabledValueType = $policy.enabledValue.ChildNodes[0].Name
                        Switch ($polEnabledValueType)
                        {
                            "string"
                            {
                                $polPossibleValues = ($polPossibleValues + "`n   " + $policy.enabledValue.string + " (Enabled)")
                            }
                            "decimal"
                            {
                                $polPossibleValues = ($polPossibleValues + "`n   " + $policy.enabledValue.decimal.value.ToString() + " (Enabled)")
                            }
                            default
                            {
                                $polPossibleValues = ""
                            }
                        }
                    }
                    If ($policy.disabledValue -ne $null)
                    {
                        $polDisabledValueType = $policy.disabledValue.ChildNodes[0].Name
                        Switch ($polDisabledValueType)
                        {
                            "string"
                            {
                                $polPossibleValues = ($polPossibleValues + "`n   " + $policy.disabledValue.string + " (Disabled)")
                            }
                            "decimal"
                            {
                                $polPossibleValues = ($polPossibleValues + "`n   " + $policy.disabledValue.decimal.value.ToString() + " (Disabled)")
                            }
                            default
                            {
                                $polPossibleValues = ""
                            }
                        }
                    }

                }
                Else
                {
                    $polValueName = ""
                    $polPossibleValues = ""
                }
                # Add a row for each policy with the basic policy information (like, displayname, explaining text, etc.) (elements will be processed on separate rows)
                    # Updated: New rows enumeration
                [void]$table.Rows.Add(
                    $file.Name,
                    $parentCategory,
                    $policy.Name,
                    $polDisplayName,
                    $policy.class,
                    $explainText,
                    $supportedOn,
                    "policy",
                    "",
                    $regkey,
                    $polValueName,
                    $polPossibleValues)

                    $rowCount = $rowCount + 1
                    $policyCount = $policyCount + 1
                #endregion


# Include a Row with policy element information (per element one row)

                #region retrieving policy element Childnode information
                $elementCount = 0
                If ($policy.elements -ne $null)
                {
                    #elements node gevonden
                    $policy.elements.ChildNodes | ForEach-Object {
                        $element = $_
                        If ($element -ne $null)
                        {
                            $elementCount = $elementCount + 1
                            $elementLabelText = ""
                            # process elements ChildNode 
                            Switch ($element.Name)
                            {
                                "#comment"                                                         # comment XML element, ignore for processing
                                {
                                    $elementType = "comment"
                                    $valueName = $element.valueName
                                    $dropdownListValues = ""
                                }

                                "list"                                                             # represents a list element in a policy, process node
                                {
                                    # Retrieve label, based on element.id and policy.name
                                    $oList = (($presentationTableChilds | Where-Object { $_.id -eq $policy.presentation.Substring(15).TrimEnd(')')}).ChildNodes | Where-Object {$_.refId -eq $element.id})

                                    # the list element has it's own registry key and different value attribute and wil lrequire a matching listBox element in the corresponding .adml file.
                                    $elementType = $oList.Name
                                    $elementLabelText = $oList.InnerText
                                    If (!([string]::IsNullOrEmpty($element.valuePrefix)))
                                    {
                                        $valueName = $element.valuePrefix + " (prefix)"
                                        $regkey = $element.key
                                        $dropdownListValues = ("Example Value keys: `n   " + $valueName.Replace(" (prefix`)","") + "1`n   " + $valueName.Replace(" (prefix)","") + "2")
                                    }
                                    Else
                                    {
                                        $valueName = "(value list)"
                                        $regkey = $element.key
                                        $dropdownListValues = ""
                                    }
                                    If ($element.additive -eq "true")
                                    {
                                        $valueName = $valueName + " (append)"
                                    }
                                }

                                "text"                                                             # represents a text element in a policy, process node
                                {                                                                  
                                    # Retrieve label, based on element.id and policy.name for the specific element
                                    $oText = (($presentationTableChilds | Where-Object { $_.id -eq $policy.presentation.Substring(15).TrimEnd(')')}).ChildNodes | Where-Object {$_.refId -eq $element.id})
                                    # the text element is generally associated with either a textbox or a combobox presentation element in the corresponding .adml file
                                    $elementType = $oText.Name
                                    $elementLabelText = $oText.label
                                    $valueName = $element.valueName

                                    If ($elementType -eq "textBox")
                                    {
                                        $dropdownListValues = ""
                                    }

                                    # the ComboBox has childelements that needs to be processes to get the suggestion values as possible values
                                    If (($elementType -eq "comboBox") -and (!([string]::IsNullOrEmpty($oText.suggestion))))
                                    {
                                        Write-Debug ("ComboBox ChildNodes= " + $oText.ChildNodes.Count)
                                        $dropdownListValues = "Suggestions:"
                                        For ($i=0; $i -le ($oText.ChildNodes.Count -1); $i++)
                                        {
                                            If ($oText.ChildNodes.Item($i).Name -eq "suggestion")  # skip the label childnode
                                            {
                                                $dropdownListValues = ($dropdownListValues + "`n   " + $oText.ChildNodes.Item($i).InnerXML)
                                            }
                                        }
                                        $valueName = $valueName + " (comboBox)"
                                    }

                                    If (($element.maxLength -ne $null) -and ($element.Name -eq "textBox"))
                                    {
                                        $dropdownListValues = ($dropdownListValues + "`n" + $element.maxLength + " (max length)")
                                    }
                                    If ($element.required -eq "true")
                                    {
                                        $valueName = $valueName + " (required)"
                                    }
                                    If ($element.soft -eq "true")
                                    {
                                        $valueName = $valueName + " (overwrite)"
                                    }
                                    If ($element.expandable -eq "true")
                                    {
                                        $valueName = $valueName + " (expandable)"
                                    }
                                }

                                "enum"                                                             # represents a enumeration element, process node
                                {
                                    $elementType = "dropdownList"

                                    # Retrieve label, based on element.id and policy.name
                                    $oEnum = (($presentationTableChilds | Where-Object { $_.id -eq $policy.presentation.Substring(15).TrimEnd(')')}).ChildNodes | Where-Object { $_.refId -eq $element.id})

                                    $elementType = $oEnum.Name
                                    $elementLabelText = $oEnum.InnerText
                                                                
                                    # Retrieving the possible items from the dropdownlist
                                    $dropdownListValues = "List items:"
                                    $itemCount = 0
                                    $element.ChildNodes | ForEach-Object {
                                        $item = $_                                                 # represents a set of display names with one value or a set of registry subkey values
                                        If (($item -ne $null) -and ($item.name -ne "#comment"))
                                        {
                                            $itemCount = $itemCount + 1
                                            $itemLabelText = ($stringTableChilds | Where-Object { $_.id -eq $item.displayName.SubString(9).TrimEnd(')') }).InnerText
                                            If ($item.value.string -ne $null)
                                            {
                                                $dropdownListValues = ($dropdownListValues + " `n     """ + $item.value.string + """ = """ + $itemLabelText + """")
                                            }
                                            If ($item.value.decimal -ne $null)
                                            {
                                                $dropdownListValues = ($dropdownListValues + " `n     """ + $item.value.decimal.value + """ = """ + $itemLabelText + """")
                                            }
                                        }
                                    }
                                    # adding the dropdownlist items to the valueName 
                                    $valueName = $element.valueName
                                    If ($element.required -eq "true")
                                    {
                                        $valueName = $valueName + " (required)"
                                    }
                                    Write-Debug "$itemCount items processed"
                                }

                                "boolean"                                                          # respresents a choice element, process node
                                {
                                    # Retrieve label, based on element.id and policy.name
                                    # require a matching checkBox element with matching id in the corresponding .adml file
                                    $oBoolean = (($presentationTableChilds | Where-Object { $_.id -eq $policy.presentation.Substring(15).TrimEnd(')')}).ChildNodes | Where-Object {$_.refId -eq $element.id})
                                    $elementType = $oBoolean.Name
                                    $elementLabelText = $oBoolean.InnerText
                                    $valueName = $element.valueName
                                    If ($oBoolean.defaultChecked -eq "true")
                                    {
                                        $valueName = $valueName + " (default checked)"
                                    }
                                    # the boolean has childelements (trueValue, falseValue, trueList, falseList)
                                    $dropdownListValues = ""
                                    If ($element.ChildNodes.Count -gt 0)
                                    {        
                                        $dropdownListValues = "Checkbox values: "
                                        For ($i=0; $i -le ($element.ChildNodes.Count -1); $i++)
                                        {
                                            If ($element.ChildNodes.Item($i).Name -eq "trueValue")
                                            {
                                                If ($element.ChildNodes.Item($i).ChildNodes[0].Name -eq "string")
                                                {
                                                    $dropdownListValues = ($dropdownListValues + "`n   " + $element.ChildNodes.Item($i).ChildNodes[0].InnerXML + " (true)")
                                                }
                                                If ($element.ChildNodes.Item($i).ChildNodes[0].Name -eq "decimal")
                                                {
                                                    $dropdownListValues = ($dropdownListValues + "`n   " + $element.ChildNodes.Item($i).ChildNodes[0].value.ToString() + " (true)")
                                                }
                                            }
                                            If ($element.ChildNodes.Item($i).Name -eq "falseValue")
                                            {
                                                If ($element.ChildNodes.Item($i).ChildNodes[0].Name -eq "string")
                                                {
                                                    $dropdownListValues = ($dropdownListValues + "`n   " + $element.ChildNodes.Item($i).ChildNodes[0].InnerXML + " (false)")
                                                }
                                                If ($element.ChildNodes.Item($i).ChildNodes[0].Name -eq "decimal")
                                                {
                                                    $dropdownListValues = ($dropdownListValues + "`n   " + $element.ChildNodes.Item($i).ChildNodes[0].value.ToString() + " (false)")
                                                }
                                            }
                                        }

                                    }
                                }

                                "decimal"                                                          # represents a number/decimal element, process node
                                {
                                    $elementType = "decimalTextBox"
                                    # Retrieve label, based on element.id and policy.name
                                    $oDecimal = (($presentationTableChilds | Where-Object { $_.id -eq $policy.presentation.Substring(15).TrimEnd(')')}).ChildNodes | Where-Object {$_.refId -eq $element.id})
                                    $elementType = $oDecimal.Name
                                    $elementLabelText = $oDecimal.InnerText
                                    $valueName = $element.valueName
                                    $dropdownListValues = "Textbox values: "
                                    If ($element.default -ne $null)
                                    {
                                        $dropdownListValues = ($dropdownListValues + "`n   "+ $element.default + " (default)")
                                    }
                                    If ($element.minvalue -ne $null)
                                    {
                                        $dropdownListValues = ($dropdownListValues + "`n   "+ $element.minValue + " (minimum)")
                                    }
                                    If ($element.maxvalue -ne $null)
                                    {
                                        $dropdownListValues = ($dropdownListValues + "`n   "+ $element.maxValue + " (maximum)")
                                    }
                                }

                                "multiText"                                                        # new element type found in ShareFileDriveMapper.admx (see W2K12 references)
                                {
                                    $elementType = "multiTextBox"
                                    # Retrieve label, based on element.id and policy.name
                                    $oMultiText = (($presentationTableChilds | Where-Object { $_.id -eq $policy.presentation.Substring(15).TrimEnd(')')}).ChildNodes | Where-Object {$_.refId -eq $element.id})
                                    $elementType = $oMultiText.Name
                                    $elementLabelText = $oMultiText.InnerText
                                    $valueName = $element.valueName
                                    $dropdownListValues = ""
                                    If ($element.required -eq "true")
                                    {
                                        $valueName = $valueName + " (required)"
                                    }
                                    If ($element.maxlength -ne $null)
                                    {
                                        $dropdownListValues = $dropdownListValues + $element.maxLength + " (max length)`n" 
                                    }
                                    If ($element.maxStrings -ne $null)
                                    {
                                        $dropdownListValues = $dropdownListValues + $element.maxStrings + " (max strings)`n"
                                    }
                                }

                                default
                                {
                                    # unknown element
                                    $elementType = "unknown"
                                    $elementLabelText = "unknown label"
                                    $valueName = $element.valueName
                                    $dropdownListValues = ""
                                }
                            }
                        }
                        Else
                        {
                            # Including the basic policy (level) setting to the table
                            $elementType="policy setting"
                            $valueName = $policy.ValueName
                            $dropdownListValues = ""
                        }
                        Write-Debug "elementType is ""$elementType"", value is $valueName"
                        If (($policy.presentation -eq "") -or ($policy.presentation -eq $null))
                        {
                            $policyText = ""
                        }
                        Else
                        {
                            $policyText = " (policy """ + $policy.presentation.Substring(15).TrimEnd(')') + """, element """  + $element.id + """) "
                        }
                        Write-Debug ("Label $elementCount : " + $elementLabelText + $policyText)

                        If ($elementType -ne "comment")
                        {
                            # Updated: New rows enumeration
                            [void]$table.Rows.Add(
                                $file.Name,
                                $parentCategory,
                                $policy.Name,
                                $polDisplayName,
                                $policy.class,
                                $explainText,
                                $supportedOn,
                                $elementType,
                                $elementLabelText,
                                $regkey,
                                $valueName,
                                $dropdownListValues)

                                $rowCount = $rowCount + 1
                        }
                    }
                }
                #endregion retrieving policy element information

                # elements are processed
                Write-Verbose ($elementCount.ToString() + " element(s) processed")

# Include a Row with policy enabledList information

                #region retrieving policy enabledList Childnode information
                $enabledListCount = 0
                If ($policy.enabledList -ne $null)
                {
                    # enabledList node found
                    Write-Debug "enabledList node found."
                    $policy.enabledList.ChildNodes | ForEach-Object {
                        $enabledItem = $_
                        If ($enabledItem -ne $null)
                        {
                            $enabledListCount = $enabledListCount + 1
                            $enabledItemType = $enabledItem.value.ChildNodes[0].Name
                            $elementType = ("enabledList " + $enabledItemType)
                            Switch ($enabledItemType)
                            {
                                "string"
                                {
                                    $dropdownListValues = ("value: " + $enabledItem.value.string)
                                }
                                "decimal"
                                {
                                    $dropdownListValues = ("value: " + $enabledItem.value.decimal.value.ToString())
                                }
                                default
                                {
                                    $dropdownListValues = ""
                                }
                            }
                            $regkey = $enabledItem.key
                            $valueName = $enabledItem.valueName

                            # Updated: New rows enumeration
                            [void]$table.Rows.Add(
                                $file.Name,
                                $parentCategory,
                                $policy.Name,
                                $polDisplayName,
                                $policy.class,
                                $explainText,
                                $supportedOn,
                                $elementType,
                                "",
                                $regkey,
                                $valueName,
                                $dropdownListValues)

                                $rowCount = $rowCount + 1
                        }
                        Else
                        {
                            # No enabledList items to process
                            $elementType="no enabledList items for policy"
                            $valueName = $policy.ValueName
                        }
                        Write-Debug "enabledListItem is ""$elementType"", value is $valueName"
                    }
                }
                #endregion retrieving policy enabledList information

# Include a Row with disabledList information

                #region retrieving policy disabledList Childnode information
                $disabledListCount = 0
                If ($policy.disabledList -ne $null)
                {
                    # disabledList node found
                    Write-Debug "disabledList node found."
                    $policy.disabledList.ChildNodes | ForEach-Object {
                        $disabledItem = $_
                        If (($disabledItem -ne $null) -and ($disabledItem.Name -ne "#comment"))
                        {
                            $disabledListCount = $disabledListCount + 1
                            $disabledItemType = $disabledItem.value.ChildNodes[0].Name
                            $elementType = ("disabledList " + $disabledItemType)
                            Switch ($disabledItemType)
                            {
                                "string"
                                {
                                    $dropdownListValues = ("value: " + $disabledItem.value.string)
                                }
                                "decimal"
                                {
                                    $dropdownListValues = ("value: " + $disabledItem.value.decimal.value)
                                }
                                default
                                {
                                    $dropdownListValues = ""
                                }
                            }
                            $regkey = $disabledItem.key
                            $valueName = $disabledItem.valueName

                            # Updated: New rows enumeration
                            [void]$table.Rows.Add(
                                $file.Name,
                                $parentCategory,
                                $policy.Name,
                                $polDisplayName,
                                $policy.class,
                                $explainText,
                                $supportedOn,
                                $elementType,
                                "",
                                $regkey,
                                $valueName,
                                $dropdownListValues)
                            
                                $rowCount = $rowCount + 1
                        }
                        Write-Debug "disabledListItem is ""$elementType"", value is $valueName"
                    }
                }
                #endregion retrieving policy disabledList information

                # counter to track the total amount of processed policies in the ADMX files
                $policyCount = $policyCount + 1
            }
            Else
            {
                Write-Debug ("Comment policies ChildNode found, node NOT processed")
            }
        }
    }
        Write-Output ("A total of " + $policyCount.ToString() + " policy settings were processed in file " + $file.Name)
}

Write-Output ("=> A total of " + $rowCount + " policy settings were translated into the CSV file.")
# Updated: Exporting the data with the right Ecoding (see XML file header: <?xml version="1.0" encoding="utf-8"?> and using the Locale Delimiter for the CSV file)
$table | Export-Csv $OutputCSVFile -NoTypeInformation -Encoding UTF8 -UseCulture -Force
Write-Output ("===> Results were saved in """ + $OutputCSVFile + """")

#endregion Execution