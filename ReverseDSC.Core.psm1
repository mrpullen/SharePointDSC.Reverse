$Global:CredsRepo = @{}

<## This function receives the path to a DSC module, and a parameter name. It then returns the type associated with the parameter (int, string, etc.). #>
function Get-DSCParamType
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)] [System.String] $ModulePath,
        [parameter(Mandatory = $true)] [System.String] $ParamName
    )

    $tokens = $null 
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ModulePath, [ref] $tokens, [ref] $errors)
    $functions = $ast.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)
    
    $functions | ForEach-Object {

        if ($_.Name -eq "Get-TargetResource") 
        {
            $function = $_
            $functionAst = [System.Management.Automation.Language.Parser]::ParseInput($_.Body, [ref] $tokens, [ref] $errors)

            $parameters = $functionAst.FindAll( {$args[0] -is [System.Management.Automation.Language.ParameterAst]}, $true)
            $parameters | ForEach-Object {
                if($_.Name.Extent.Text -eq $ParamName)
                {
                    $attributes = $_.Attributes
                    $attributes | ForEach-Object{
                        if($_.TypeName.FullName -like "System.*")
                        {
                            return $_.TypeName.FullName
                        }
                        elseif($_.TypeName.FullName.ToLower() -eq "microsoft.management.infrastructure.ciminstance")
                        {
                            return "System.Collections.Hashtable"
                        }
                        elseif($_.TypeName.FullName.ToLower() -eq "string")
                        {
                            return "System.String"
                        }
                        elseif($_.TypeName.FullName.ToLower() -eq "boolean")
                        {
                            return "System.Boolean"
                        }
                        elseif($_.TypeName.FullName.ToLower() -eq "string[]")
                        {
                            return "System.String[]"
                        }
                    }
                }
            }
        }
     }
     return $null
 }

<## This function loops through a HashTable and returns a string that combines all the Key/Value pairs into a DSC param block. #>
function Get-DSCBlock
{
    [CmdletBinding()]
    param(        
        [System.String] $ModulePath,
        [System.Collections.Hashtable] $Params
    )

    $dscBlock = ""
    $foundInstallAccount = $false
    $Params.Keys | % { 
        $paramName = $_
        $paramType = Get-DSCParamType -ModulePath $ModulePath -ParamName "`$$_"

        $value = $null
        if($paramType -eq "System.String")
        {
            if(!$null -eq $Params.Item($_))
            {
                $value = "`"" + $Params.Item($_).ToString().Replace("`"", "```"") + "`""
            }
            else
            {
                $value = "`"" + $Params.Item($_) + "`""
            }
        }
        elseif($paramType -eq "System.Boolean")
        {
            $value = "`$" + $Params.Item($_)
        }
        elseif($paramType -eq "System.Management.Automation.PSCredential" -and $_ -ne "InstallAccount")
        {
            $value = "`$CredsFarmAccount #`"" + ($Params.Item($_)).Username + "`""
        }
        elseif($paramType -eq "System.Collections.Hashtable")
        {
            $value = "@{"
            $hash = $Params.Item($_)
            $hash.Keys | % {
                try
                {
                    $value += $_ + " = `"" + $hash.Item($_) + "`"; "
                    $value += "}"
                }
                catch
                {
                    $value = $hash
                }
            }
            
        }
        elseif($paramType -eq "System.String[]")
        {
            $value = "@("
            $hash = $Params.Item($_)
            $hash| % {
                $value += "`"" + $_ + "`","
            }
            if($value.Length -gt 2)
            {
                $value = $value.Substring(0,$value.Length -1)
            }
            $value += ")"
        }
        else
        {
            if($_ -eq "InstallAccount")
            {
                $value = "`$CredsFarmAccount"
                $foundInstallAccount = $true
            }
            else
            {
                $value = $Params.Item($_)
            }
        }
        $dscBlock += "            " + $_  + " = " + $value + ";`r`n"
    }

    if(!$foundInstallAccount)
    {
        $dscBlock += "            PsDscRunAsCredential=`$CredsFarmAccount;`r`n"
    }
    
    return $dscBlock
}

<## This function generates an empty hash containing fakes values for all input parameters of a Get-TargetResource function. #>
function Get-DSCFakeParameters{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)] [System.String] $ModulePath
    )

    $params = @{}

    $tokens = $null 
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ModulePath, [ref] $tokens, [ref] $errors)
    $functions = $ast.FindAll( {$args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)
    
    $functions | ForEach-Object {

        if ($_.Name -eq "Get-TargetResource") 
        {
            $function = $_
            $functionAst = [System.Management.Automation.Language.Parser]::ParseInput($_.Body, [ref] $tokens, [ref] $errors)

            $parameters = $functionAst.FindAll( {$args[0] -is [System.Management.Automation.Language.ParameterAst]}, $true)
            $parameters | ForEach-Object {   
                $paramName = $_.Name.Extent.Text
                $attributes = $_.Attributes
                $found = $false

                <# Loop once to figure out if there is a validate Set to use. #>
                $attributes | ForEach-Object{
                    if($_.TypeName.FullName -eq "ValidateSet")
                    {
                        $params.Add($paramName.Replace("`$", ""), $_.PositionalArguments[0].ToString().Replace("`"", "").Replace("'",""))
                        $found = $true
                    }
                }
                $attributes | ForEach-Object{
                    if(!$found)
                    {
                        if($_.TypeName.FullName -eq "System.String" -or $_.TypeName.FullName -eq "String")
                        {
                            $params.Add($paramName.Replace("`$", ""), "*")
                            $found = $true
                        }
                        elseif($_.TypeName.FullName -eq "System.UInt32")
                        {
                            $params.Add($paramName.Replace("`$", ""), 0)
                            $found = $true
                        }
                        elseif($_.TypeName.FullName -eq "System.Management.Automation.PSCredential")
                        {
                            $params.Add($paramName.Replace("`$", ""), $null)                            
                            $found = $true
                        }
                        elseif($_.TypeName.FullName -eq "System.Management.Automation.Boolean" -or $_.TypeName.FullName -eq "System.Boolean" -or $_.TypeName.FullName -eq "Boolean")
                        {
                            $params.Add($paramName.Replace("`$", ""), $true)
                            $found = $true
                        }
                        elseif($_.TypeName.FullName -eq "System.String[]" -or $_.TypeName.FullName -eq "String[]")
                        {
                            $params.Add($paramName.Replace("`$", ""),[string]@("1","2"))
                            $found = $true
                        }
                    }
                }
            }
        }
     }
     return $params
}

function Export-TargetResource()
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)] [System.String] $ResourceName,
        [parameter(Mandatory = $true)] [System.Collections.Hashtable] $MandatoryParameters,
        [parameter(Mandatory = $false)] [System.String] $DependsOnClause
    )
    $ModulePath = (Get-DscResource $ResourceName | select Path).Path.ToString()
    $friendlyName = Get-ResourceFriendlyName -ModulePath $ModulePath
    $fakeParameters = Get-DSCFakeParameters -ModulePath $ModulePath

    <# Nik20170109 - Replace each Fake Parameter by the ones received as function arguments #>
    $finalParams = @{}
    foreach($fakeParameter in $fakeParameters.Keys)
    {
        if($MandatoryParameters.ContainsKey($fakeParameter))
        {
            $finalParams.Add($fakeParameter,$MandatoryParameters.Get_Item($fakeParameter))
        }
    }

    Import-Module $ModulePath
    $results = Get-TargetResource @finalParams
    
    $exportContent = "        " + $friendlyName + " " + [System.Guid]::NewGuid().ToString() + "`r`n"
    $exportContent += "        {`r`n"
    $exportContent += Get-DSCBlock -ModulePath $ModulePath -Params $results
    if($null -ne $DependsOnClause)
    {
        $exportContent += $DependsOnClause
    }
    $exportContent += "        }`r`n"
    return $exportContent
}

function Get-ResourceFriendlyName()
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $true)] [System.String] $ModulePath
    )

    $tokens = $null 
    $errors = $null
    $schemaPath = $ModulePath.Replace(".psm1", ".schema.mof")
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($schemaPath, [ref] $tokens, [ref] $errors)

    for($i = 0; $i -lt $tokens.Length; $i++)
    {
        if($tokens[$i].Text.ToLower() -eq "friendlyname" -and ($i+2) -le $tokens.Length)
        {
            return $tokens[$i+2].Text.Replace("`"", "")
        }
    }
    return $null
}

<# Region Helper Methods #>
<## This function receives a user name and returns the "Display Name" for that user. This function is primarly used to identify the Farm (System) account. #>
function Get-Credentials([string] $userName)
{
    if($Global:CredsRepo.Contains($userName.ToLower()))
    {
        return $Global:CredsRepo[$userName]
    }
    else
    {
        $creds = Get-Credential -Message "Please Provide Credentials for $userName" -UserName $userName
        $Global:CredsRepo.Add($userName.ToLower(), $creds)
        return $creds
    }
}

function Resolve-Credentials([string] $userName)
{
    if($userName -eq $Script:spCentralAdmin.ApplicationPool.ProcessAccount.Name)
    {
        return "`$CredsFarmAccount"
    }
    else
    {
        $userNameParts = $userName.Split('\')
        if($userNameParts.Length -gt 1)
        {
            return "`$Creds" + $userNameParts[1]
        }
        return "`$Creds" + $userName
    }
    return $userName
}

function Save-Credentials([System.Management.Automation.PSCredential] $creds)
{
    if($Global:CredsRepo.Contains($creds.UserName.ToLower()))
    {
        return $true
    }
    else
    {
        $Global:CredsRepo.Add($creds.UserName.ToLower(), $creds)
        return $false
    }
}