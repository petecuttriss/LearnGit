CD "C:\Users\il.migrate\Documents"
. .\SQLFunctions.ps1
.\PrepKaitunaContractors.ps1

$networkSharePath = "\\pradfs2\Ferret\KaitunaContractorsDocs\Files000"
$dataSource = "FerretMigrate.dbo.vMetaData_KaitunaContractors_1"
$dataResults = "FerretMigrate.dbo.tMigrate_KaitunaContractors"

$connection = ConnectToSQL -serverName PRADSQL2 -databaseName FerretMigrate

$count = 0
$nomapping = 0

foreach($accountFolder in Get-ChildItem -Path $networkSharePath -Recurse | ?{$_.PSIsContainer})
{
    Write-Host "Processing account No." $accountFolder.Name -ForegroundColor Cyan

    foreach($file in Get-ChildItem -Path $accountFolder.PSPath)
    {
        $dataTable = ""
        $dataTable = QuerySQL -connection $connection -dataSource $dataSource -accountFolder $accountFolder -fileName $file
        $row = $dataTable[$dataTable.Count - 1]

        if($row.MigrationStatus -ne "File Exists")
        {
            Write-Host $file.FullName ": " -NoNewline

            if($row.AccountType -eq "SUPPLIER CONTRACTS")
            {
                $siteurl = "https://nelsonmanagement.sharepoint.com/sites/services"
                $libraryName = "Active"
            }
            elseif($row.AccountType -eq "DEBTOR CONTRACTS & CORRESPONDENCE")
            {
                $siteurl = "https://nelsonmanagement.sharepoint.com/sites/contracts"
                $libraryName = "Active"
            }
            else
            {
                Write-Host "No match for AccountType" $row.DocID -ForegroundColor Red
            }

            # Library to migrate to and content types to use
            $newWeb = Get-PnPWeb
            if(!$connected -or $newWeb.Url -ne $siteurl)
            {
                if(!$cred){$cred = Get-Credential}
                Connect-PnPOnline -Url $siteurl -Credentials $cred
                $connected = $true
                $newWeb = Get-PnPWeb
            }
            
            # Library to migrate to
            $newList = Get-PnPList -Identity $libraryName
            $eDocumentCT = Get-PnPContentType -List $newList | ?{$_.Name -eq "eDocument"}
            $eMailCT = Get-PnPContentType -List $newList | ?{$_.Name -eq "eMail"}

            if($row.FileName -like "*.msg")
            {
                $contentTypeID = $eMailCT.Id.StringValue
            }
            else
            {
                $contentTypeID = $eDocumentCT.Id.StringValue
            }
        
            if($row.DocID)
            {
                Write-Host $row.DocID ":" $row.Destination ": " -NoNewline
                $folderPath = ($row.Destination).Trim()

                $fileExists = Get-PnPFile -Url ($folderPath + "/" + $file.Name) -ErrorAction SilentlyContinue
                if(!$fileExists)
                {
                    $tempFile = "C:\Users\il.migrate\Documents\" + $file.Name
                    Copy-Item -Path $file.FullName -Destination $tempFile -Force
                    Set-ItemProperty -Path $tempFile -Name IsReadOnly -Value $false
                    $fileStream = New-Object IO.FileStream($tempFile,[System.IO.FileMode]::Open)
                    $uploadStatus = Add-PnPFile -Stream $fileStream -FileName $file.Name.Replace("%","Percent") -Folder $folderPath `
                        -Values @{ `
                        Modified=$row.DocDate; Created=$row.DocDate; `
                        ContentTypeId=$contentTypeID; `
                        Title=$row.Description; `
                        Description1=$row.Description; `
                        DocTypeName=$row.DocTypeName; `
                        DocumentType="CONTRACT, Variation, Agreement";`
                        To=$row.DocEmailTo; `
                        From1=$row.DocEmailFrom; `
                        AccountNumber=$row.AccountNumber; `
                        AccountName=$row.AccountName; `
                        AccountType=$row.AccountType; `
                        FileName=$row.FileName; `
                        AccountField3=$row.AccountField3; `
                        AccountField3Label=$row.AccountField3Label
                    }

                    if($uploadStatus.UniqueId)
                    {
                        Write-Host "Uploaded" -ForegroundColor Green
                        $migrationStatus = "Completed"
                        UpdateMigrationStatus -connection $connection -dataSource $dataResults -docID $row.DocID -migrationStatus $migrationStatus
                    }
                    else
                    {
                        Write-Host "ERROR" -ForegroundColor Red
                        $migrationStatus = "ERROR - Failed to upload - " + $folderPath + "/" + $file.Name
                        UpdateMigrationStatus -connection $connection -dataSource $dataResults -docID $row.DocID -migrationStatus $migrationStatus
                    }
                }
                else
                {
                    Write-Host "File Exists" -ForegroundColor Gray
                    $migrationStatus = "File Exists"
                    UpdateMigrationStatus -connection $connection -dataSource $dataResults -docID $row.DocID -migrationStatus $migrationStatus
                }
                $count++
            }
            else
            {
                $accountDestination = QueryAccountDestination -connection $connection -dataSource $dataSource -accountNumber $accountFolder.Name
                $accountDestinationRow = $accountDestination[$accountDestination.Count - 1]
                $folderPath = $accountDestinationRow.Destination
                if(!$folderPath)
                {
                    $folderPath = $libraryName + "/Orphaned Files"
                    $folder = Get-PnPFolder -RelativeUrl $folderPath -ErrorAction SilentlyContinue
                    if(!$folder)
                    {
                        Write-Host "Create new folder at" $folderPath -ForegroundColor Cyan
                        $folder = Add-PnPFolder -Name "Orphaned Files" -Folder $libraryName
                    }
                    
                    $fileExists = Get-PnPFile -Url ($folderPath + "/" + $file.Name) -ErrorAction SilentlyContinue
                    if(!$fileExists)
                    {
                        Write-Host "No mapping in Ferret" -ForegroundColor Yellow
                        $tempFile = "C:\Users\il.migrate\Documents\" + $file.Name
                        Copy-Item -Path $file.FullName -Destination $tempFile -Force
                        Set-ItemProperty -Path $tempFile -Name IsReadOnly -Value $false
                        $fileStream = New-Object IO.FileStream($tempFile,[System.IO.FileMode]::Open)
                        Add-PnPFile -Stream $fileStream -FileName $file.Name -Folder $folderPath `
                            -Values @{ `
                            Modified=$file.LastWriteTime; Created=$file.LastWriteTime; `
                            ContentTypeId=$contentTypeID; `
                            Title=$file.Name; `
                            DocumentType="CONTRACT, Variation, Agreement";`
                            FileName=$file.Name; `
                        } | Out-Null
                    }
                }
                else
                {
                    $folderPath = $folderPath.Trim()
                    $fileExists = Get-PnPFile -Url ($folderPath + "/" + $file.Name) -ErrorAction SilentlyContinue
                    if(!$fileExists)
                    {
                        Write-Host "No mapping in Ferret" -ForegroundColor Yellow
                        $tempFile = "C:\Users\il.migrate\Documents\" + $file.Name
                        Copy-Item -Path $file.FullName -Destination $tempFile -Force
                        Set-ItemProperty -Path $tempFile -Name IsReadOnly -Value $false
                        $fileStream = New-Object IO.FileStream($tempFile,[System.IO.FileMode]::Open)
                        Add-PnPFile -Stream $fileStream -FileName $file.Name -Folder $folderPath `
                            -Values @{ `
                            Modified=$file.LastWriteTime; Created=$file.LastWriteTime; `
                            ContentTypeId=$contentTypeID; `
                            Title=$file.Name; `
                            DocumentType="CONTRACT, Variation, Agreement";`
                            AccountNumber=$accountDestinationRow.AccountNumber; `
                            AccountName=$accountDestinationRow.AccountName; `
                            AccountType=$accountDestinationRow.AccountType; `
                            FileName=$file.Name; `
                            AccountField3=$accountDestinationRow.AccountField3; `
                            AccountField3Label=$accountDestinationRow.AccountField3Label
                        } | Out-Null
                    }
                }

                $nomapping++
            }
            if($fileStream)
            {
                $fileStream.Close()
                $fileStream = ""
                Remove-Item $tempFile
            }
        }
    }

    <#if($count -ge 10)
    {
        break
    }#>
}

Write-Host "Uploaded count:" $count
Write-Host "No mapping:" $nomapping

CloseSQLConnection -connection $connection