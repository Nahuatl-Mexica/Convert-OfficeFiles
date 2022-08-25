Function ConvertDOC-ToDOCX {
[cmdletbinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)]
        [string[]]$Path,

        [Parameter(Mandatory=$false)]
        [string]$Destination,

        [Parameter(Mandatory=$false)]
        [int32]$Delay = 500,

        [Parameter(Mandatory=$false)]
        [switch]$RemoveOld
    )
    DynamicParam 
    {
        if (-not[System.IO.Path]::HasExtension($Path)) {
            $recurseAttribute = New-Object System.Management.Automation.ParameterAttribute
            $recurseAttribute.Mandatory = $false
            $attributeCollection = new-object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($recurseAttribute)
            $recurseParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Recurse', [switch], $attributeCollection)
            $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Recurse', $recurseParam)
            return $paramDictionary
       }
    } 
    Begin 
    {
        [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word") 
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $False
        $saveAs = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
        $wordApp.DisplayAlerts = "wdAlertsNone"
        $wordApp.WordBasic.DisableAutoMacros
        $wordApp.AutomationSecurity = "msoAutomationSecurityForceDisable"

        $filterType = "*.doc"

        $conversionCall = {
            if ($PSCmdlet.ShouldProcess($File.BaseName, "Converting File from DOC to DOCX")) {
                Write-Verbose -Message "Converting $($File.FullName)"
                Write-Verbose -Message "Opening file $($File.BaseName) using MSWord COM Object."
			    $openDoc  = $wordApp.Documents.Open($File.FullName)
			    $saveName = ($File.FullName).Substring(0, ($File.FullName).LastIndexOf("."))
                    if ($Destination) {
                        Write-Verbose -Message "Destination was specfied: $Destination"
                        $fileName = Split-Path -Path $saveName -Leaf
                        $saveName = Join-Path -Path $Destination -ChildPath $fileName
                    }

                Write-Verbose -Message "Performing conversion. . ."
			    $openDoc.Convert()

                Write-Verbose -Message "Saving file as DOCX."
			    $openDoc.SaveAs2([system.object]$saveName, $saveAs);

                Write-Verbose -Message "File saved!"
                Write-Verbose -Message "Closing file.. . Releasing word hook."
                $openDoc.close()

                Start-Sleep -Milliseconds $Delay
                (Get-Item -Path ($saveName + ".docx")).CreationTime = $File.CreationTime
                (Get-Item -Path ($saveName + ".docx")).LastWriteTime = $File.LastWriteTime
            }
            if ($RemoveOld) {
                Write-Verbose -Message "[RemovedOld] switch specified!"
                Write-Verbose -Message "Removing old Word Document."
                Remove-Item -LiteralPath $File.FullName
            }
        }
    }
    Process
    {
        Try {
            foreach ($File in $Path) 
            {
                $testPath = Get-Item -Path $File -ErrorAction SilentlyContinue
                    if ($null -ne $testPath -and $testPath -is [System.IO.FileInfo]) {
                        $File = $testPath
                        & $conversionCall
                    }
                    elseif ($null -ne $testPath -and $testPath -is [System.IO.DirectoryInfo]) {
                        if ($PSBoundParameters['Recurse'].IsPresent) {
                            $foundDocs = Get-ChildItem -LiteralPath $File -Include $filterType -File -Recurse
                                foreach ($File in $foundDocs)
                                {
                                    & $conversionCall
                                }
                        }
                        else {
                            $foundDocs = Get-ChildItem -LiteralPath $File -Include $filterType -File
                                foreach ($File in $foundDocs)
                                {
                                    & $conversionCall
                                }
                        }
                    }
                    else {
                        Throw "Path is not valid!"
                    }
            }
        }
        Catch {
            $_
        }
        Finally {
            $WordApp.quit()
	        $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApp)
            Remove-Variable -Name wordApp
        }
    }
    End 
    { 
        [gc]::collect()
	    [gc]::WaitForPendingFinalizers()
    }
}

Function ConvertXLS-ToXLSX {
[cmdletbinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)]
        [string[]]$Path,

        [Parameter(Mandatory=$false)]
        [string]$Destination,

        [Parameter(Mandatory=$false)]
        [switch]$RemoveOld
    )
    DynamicParam 
    {
        if (-not[System.IO.Path]::HasExtension($Path)) {
            $recurseAttribute = New-Object System.Management.Automation.ParameterAttribute
            $recurseAttribute.Mandatory = $false
            $attributeCollection = new-object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($recurseAttribute)
            $recurseParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Recurse', [switch], $attributeCollection)
            $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Recurse', $recurseParam)
            return $paramDictionary
       }
    } 
    Begin 
    {
        [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Excel")
		$ExcelApp = New-Object -ComObject Excel.Application
		$ExcelApp.Visible = $False
		$ExcelApp.AutomationSecurity = "msoAutomationSecurityForceDisable"
		$ExcelApp.DisplayAlerts = $False
		$ExcelUsed = $True

        $filterType = "*.xls"

        $conversionCall = {
            if ($PSCmdlet.ShouldProcess($File.BaseName, "Converting File from XLX to XLSX")) {
                Write-Verbose -Message "Converting $($File.FullName)"
                Write-Verbose -Message "Opening file $($File.BaseName) using Excel COM Object."
			    $ExcelWorkbook = $ExcelApp.WorkBooks.Open($File.Fullname, 0, $True, 5, "")
			    $saveName = ($File.FullName).Substring(0, ($File.FullName).LastIndexOf("."))
                    if ($Destination) {
                        $fileName = Split-Path -Path $saveName -Leaf
                        $saveName = Join-Path -Path $Destination -ChildPath $fileName
                    }
                    if ($ExcelWorkbook.HasVBProject) {
		                $SaveAs = [Enum]::Parse([Microsoft.Office.Interop.Excel.XlFileFormat], "xlOpenXMLWorkbookMacroEnabled")
	                }
	                else {
		                $SaveAs = [Enum]::Parse([Microsoft.Office.Interop.Excel.XlFileFormat], "xlOpenXMLWorkbook")
	                }
	            Write-Verbose -Message "Performing conversion. . ."
                Write-Verbose -Message "Saving file as XLSX."
	            $ExcelWorkbook.SaveAs([system.object]$SaveName, $SaveAs)

                Write-Verbose -Message "File saved!"
                Write-Verbose -Message "Closing file.. . Releasing Excel hook."
                Start-Sleep -Milliseconds 500
                $ExcelWorkbook.Close()

                (Get-Item -Path ($saveName + ".XLSX")).CreationTime = $File.CreationTime
                (Get-Item -Path ($saveName + ".XLSX")).LastWriteTime = $File.LastWriteTime
            }
            if ($RemoveOld) {
                Write-Verbose -Message "[RemovedOld] switch specified!"
                Write-Verbose -Message "Removing old Excel file."
                Remove-Item -LiteralPath $File.FullName
            }
        }
    }
    Process
    {
        Try {
            foreach ($File in $Path) 
            {
                $testPath = Get-Item -Path $File -ErrorAction SilentlyContinue
                    if ($null -ne $testPath -and $testPath -is [System.IO.FileInfo]) {
                        $File = $testPath
                        & $conversionCall
                    }
                    elseif ($null -ne $testPath -and $testPath -is [System.IO.DirectoryInfo]) {
                        if ($PSBoundParameters['Recurse'].IsPresent) {
                            $foundDocs = Get-ChildItem -LiteralPath $File -Include $filterType -File -Recurse
                                foreach ($File in $foundDocs)
                                {
                                    & $conversionCall
                                }
                        }
                        else {
                            $foundDocs = Get-ChildItem -LiteralPath $File -Include $filterType -File
                                foreach ($File in $foundDocs)
                                {
                                    & $conversionCall
                                }
                        }
                    }
                    else {
                        Throw "Path is not valid!"
                    }
            }
        }
        Catch {
            $_
        }
        Finally {
            $ExcelApp.Quit()
	        $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp)
            Remove-Variable -Name ExcelApp
        }
    }
    End 
    { 
        [gc]::collect()
	    [gc]::WaitForPendingFinalizers()
    }
}

Function ConvertPPT-ToPPTX {
[cmdletbinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$true)]
        [string[]]$Path,

        [Parameter(Mandatory=$false)]
        [string]$Destination,

        [Parameter(Mandatory=$false)]
        [int32]$Delay = 1600,

        [Parameter(Mandatory=$false)]
        [switch]$RemoveOld
    )
    DynamicParam 
    {
        if (-not[System.IO.Path]::HasExtension($Path)) {
            $recurseAttribute = New-Object System.Management.Automation.ParameterAttribute
            $recurseAttribute.Mandatory = $false
            $attributeCollection = new-object System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($recurseAttribute)
            $recurseParam = New-Object System.Management.Automation.RuntimeDefinedParameter('Recurse', [switch], $attributeCollection)
            $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add('Recurse', $recurseParam)
            return $paramDictionary
       }
    } 
    Begin 
    {
        [void][Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.PowerPoint")
		$PowerApp = New-Object -ComObject PowerPoint.Application
		#$PowerApp.Visible = $False
		$PowerApp.AutomationSecurity = "msoAutomationSecurityForceDisable"
		$PowerApp.DisplayAlerts = "ppAlertsNone"

        $oldExtension = "*.PPT"
        $newExtension = ".PPTX"

        $conversionCall = {
            if ($PSCmdlet.ShouldProcess($File.BaseName, "Converting File from $oldExtension to $newExtension")) {
                Write-Verbose -Message "Converting $($File.FullName)"
                Write-Verbose -Message "Opening file $($File.BaseName) using PowerPoint COM Object."
			    $ppPresentation = $PowerApp.Presentations.Open2007($File.Fullname, $false, $True, $False, $True)
			    $saveName = ($File.FullName).Substring(0, ($File.FullName).LastIndexOf("."))
                    if ($Destination) {
                        Write-Verbose -Message "Destination was specfied: $Destination"
                        $fileName = Split-Path -Path $saveName -Leaf
                        $saveName = Join-Path -Path $Destination -ChildPath $fileName
                    }
                    if ($ppPresentation.HasVBProject) {
		                $SaveAs = [Enum]::Parse([Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType], "ppSaveAsOpenXMLPresentationMacroEnabled")
	                }
	                else {
		                $SaveAs = [Enum]::Parse([Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType], "ppSaveAsOpenXMLPresentation")
	                }
	            Write-Verbose -Message "Performing conversion. . ." # this may not be needed - saving should be all you need.
                $ppPresentation.EnsureAllMediaUpgraded()

                Write-Verbose -Message "Saving file as $newExtension."
	            $ppPresentation.SaveCopyAs([system.object]$SaveName, $SaveAs)

                Write-Verbose -Message "File saved!"
                Write-Verbose -Message "Closing file.. . Releasing PowerPoint hook."
                $ppPresentation.Close() 

                Start-Sleep -Milliseconds $Delay
                (Get-Item -Path ($saveName + $newExtension)).CreationTime = $File.CreationTime
                (Get-Item -Path ($saveName + $newExtension)).LastWriteTime = $File.LastWriteTime
            }
            if ($RemoveOld) {
                Write-Verbose -Message "[RemovedOld] switch specified!"
                Write-Verbose -Message "Removing old PowerPoint file."
                Remove-Item -LiteralPath $File.FullName
            }
        }
    }
    Process
    {
        Try {
            foreach ($File in $Path) 
            {
                $testPath = Get-Item -Path $File -ErrorAction SilentlyContinue
                    if ($null -ne $testPath -and $testPath -is [System.IO.FileInfo]) {
                        $File = $testPath
                        & $conversionCall
                    }
                    elseif ($null -ne $testPath -and $testPath -is [System.IO.DirectoryInfo]) {
                        if ($PSBoundParameters['Recurse'].IsPresent) {
                            $foundDocs = Get-ChildItem -LiteralPath $File -Include $oldExtension -File -Recurse
                                foreach ($File in $foundDocs)
                                {
                                    & $conversionCall
                                }
                        }
                        else {
                            $foundDocs = Get-ChildItem -LiteralPath $File -Include $oldExtension -File
                                foreach ($File in $foundDocs)
                                {
                                    & $conversionCall
                                }
                        }
                    }
                    else {
                        Throw "Path is not valid!"
                    }
            }
        }
        Catch {
            $_
        }
        Finally {
            $PowerApp.Quit()
	        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerApp)
            Remove-Variable -Name PowerApp
        }
    }
    End 
    { 
        [gc]::collect()
	    [gc]::WaitForPendingFinalizers()
    }
}

Function Convert-OfficeFiles {
[cmdletbinding(SupportsShouldProcess=$true)]
    Param (
        [Parameter(Mandatory=$false)]
        [string[]]$Path = $(Read-Host -Prompt "Enter Path"),

        [Parameter(Mandatory=$false)]
        [string]$Destination,
        
        [Parameter(Mandatory=$false,
                   DontShow=$true)]
        [array]$Include = @('*.ppt','*.doc','*.xls')
    )
    Begin { 
        $beginMessage = "Computer Support does not guarantee the use of this script. Use at your own risk. Continue with caution."    
        [System.Windows.MessageBox]::Show($beginMessage)

        $colorSplat = @{
            ForegroundColor = 'Green'
            BackgroundColor = 'Black'
        }
    }
    Process
    {
        foreach ($Path in $Path)
        {
            try {
                if ((Test-Path -LiteralPath $Path) -and (Get-Item -Path $Path) -is [System.IO.DirectoryInfo]) {
                    $Files =  Get-ChildItem -Path "$Path\*" -Include $Include
                    if ($Files) {
                        $newFolder = Join-Path -Path $Path -ChildPath "Converted"
                            if (-not (Test-Path -LiteralPath $newFolder)) {
                                $Destination = (New-Item -Path $newFolder -ItemType Directory).FullName
                            }
                            elseif (-not$PSBoundParameters.ContainsKey('Destination') -and (Test-Path -LiteralPath $newFolder)) {
                                $Destination = $newFolder
                            }
                        switch ($Files) 
                        {
                            {$_.Name -match "\.doc$"} { 
                                Write-Host -Object "`nFound Word Document: $($_.Name)"
                                Write-Host -Object "Converting File!"
                                ConvertDOC-ToDOCX -Path $_.FullName -ErrorAction Stop -Destination $Destination
                                Write-Host -Object "Completed!" @colorSplat
                                continue
                            }
                            {$_.Name -match "\.xls$"} { 
                                Write-Host -Object "`nFound Excel Document: $($_.Name)"
                                Write-Host -Object "Converting File!"
                                ConvertXLS-ToXLSX -Path $_.FullName -ErrorAction Stop -Destination $Destination
                                Write-Host -Object "Completed!" @colorSplat
                                continue
                            }
                            {$_.Name -match "\.ppt$"} { 
                                Write-Host -Object "`nFound PowerPoint Document: $($_.Name)"
                                Write-Host -Object "Converting File!"
                                ConvertPPT-ToPPTX -Path $_.FullName -ErrorAction Stop -Destination $Destination
                                Write-Host -Object "Completed!" @colorSplat
                                continue
                            }
                        }
                    }
                    else {
                        Write-Verbose -Message "No old Office files found in $Path"
                    }
                }
                else {
                    throw "[$Path] is Invalid!"
                }
            }
            catch {
                Write-Host -Object $_.Exception.Message -ForegroundColor Red -BackgroundColor Black
            }
        }
    }
    End { 
        $Message1 = "`t`t`t!WARNING!`n`nIt is the user's responsibility to ensure that all converted files were converted. Computer Support is not liable for missing, or corrupted files. `n`n`n"
        $Message2 = "Use of this script is not guaranteed. Computer Support will not provide further support with use of this script."
        [System.Windows.MessageBox]::Show($Message1 + $Message2)
    }
}
