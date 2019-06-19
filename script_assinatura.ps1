#############################################################################
#                                     			 		    #
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 		                    #
#############################################################################
#variables
$SignatureName = 'nome que vai ser adicionado como assinatura no outlook'
$SigSource = 'caminho para a pasta de aruivos onde é pego o template da assinatura'
$SignatureVersion = "2018.9"  

$ForceSignature = '0' #'0' = editable ; '1' non-editable and forced.
 
#Environment variables
$AppData=$env:appdata
$SigPath = '\Microsoft\Signatures'
$LocalSignaturePath = $AppData+$SigPath
$RemoteSignaturePathFull = $SigSource
	
# stop outlook process
ps outlook -ErrorAction SilentlyContinue | kill -PassThru

#Copy file
If (!(Test-Path -Path $LocalSignaturePath\$SignatureVersion))
{ 
    New-Item -Path $LocalSignaturePath\$SignatureVersion -Type Directory
}
Elseif (Test-Path -Path $LocalSignaturePath\$SignatureVersion)
{
    Remove-Item -Path $LocalSignaturePath\$SignatureVersion 									#remove pasta padrão criada pelo office '2018.9'
    Write-Host "Signature already exists, Script will now exit..." -ForegroundColor Yellow				
    New-Item -Path $LocalSignaturePath\$SignatureVersion -Type Directory						#adiciona a pasta novamente, se não remover a assinatura não é adicionada ou adiciona em branco

}

#Check signature path 
if (!(Test-Path -path $LocalSignaturePath)) {
	New-Item $LocalSignaturePath -Type Directory		#checa o caminho das assinaturas
}

#Get Active Directory information for logged in user
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()


#Copy signature templates from domain to local Signature-folder
Write-Host "Copying Signatures" -ForegroundColor Green
Copy-Item ("$Sigsource"+'\*') "$LocalSignaturePath\" -Force
Copy-Item ("$Sigsource"+'\'+"$signatureName"+'_arquivos\*') ("$LocalSignaturePath"+'\'+"$SignatureName"+'_arquivos\') -Force 

#Insert variables from Active Directory to rtf signature-file
#$MSWord = New-Object -ComObject word.application
$fullPath = "$LocalSignaturePath\$SignatureName.htm"
$fullPathrtf = "$LocalSignaturePath\$SignatureName.rtf"
$fullPathtxt = "$LocalSignaturePath\$SignatureName.txt"
#$MSWord.Documents.Open($fullPath)
#$MSWord.visible = $true
#User Name $ Designation 

			#Adiciona no template os dados do AD em 3 formatos diferentes (html/rtf/txt)

#update html
(get-content $fullPath) -replace ('displayName', $ADUser.displayName) | out-file $fullPath
(get-content $fullPath) -replace ('title', $ADUser.title) | out-file $fullPath
(get-content $fullPath) -replace ('TelephoneNumber', $ADUser.telephoneNumber) | out-file $fullPath
(get-content $fullPath) -replace ('homePhone', $ADUser.homePhone) | out-file $fullPath
(get-content $fullPath) -replace ('MobileNumber1', $ADUser.mobile) | out-file $fullPath
(get-content $fullPath) -replace ('mail', $ADUser.mail) | out-file $fullPath

#updatertf
(get-content $fullPathrtf) -replace ('displayName', $ADUser.displayName) | out-file $fullPathrtf -Encoding default
(get-content $fullPathrtf) -replace ('title', $ADUser.title) | out-file $fullPathrtf -Encoding default
(get-content $fullPathrtf) -replace ('TelephoneNumber', $ADUser.telephoneNumber) | out-file $fullPathrtf -Encoding default
(get-content $fullPathrtf) -replace ('homePhone', $ADUser.homePhone) | out-file $fullPathrtf -Encoding default
(get-content $fullPathrtf) -replace ('MobileNumber1', $ADUser.mobile) | out-file $fullPathrtf -Encoding default
(get-content $fullPathrtf) -replace ('mail', $ADUser.mail) | out-file $fullPathrtf -Encoding default

#updatetxt
(get-content $fullPathtxt) -replace ('displayName', $ADUser.displayName) | out-file $fullPathtxt
(get-content $fullPathtxt) -replace ('title', $ADUser.title) | out-file $fullPathtxt
(get-content $fullPathtxt) -replace ('TelephoneNumber', $ADUser.telephoneNumber) | out-file $fullPathtxt
(get-content $fullPathtxt) -replace ('homePhone', $ADUser.homePhone) | out-file $fullPathtxt
(get-content $fullPathtxt) -replace ('MobileNumber1', $ADUser.mobile) | out-file $fullPathtxt
(get-content $fullPathtxt) -replace ('mail', $ADUser.mail) | out-file $fullPathtxt

#Set Signature as default

#Enforce embedded pictures in outlook
#if (!(Test-Path -Path C:\ProgramData\Microsoft\Office\16.0\Outlook\Options\Mail)) { New-Item -Path C:\ProgramData\Microsoft\Office\16.0\Outlook\Options\Mail -ItemType Directory -Force }
#New-ItemProperty C:\ProgramData\Microsoft\Office\Outlook\Options\Mail -Name 'Send Pictures With Document' -Value 1 -PropertyType 4 -Force

if (!(Test-Path -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail)) { New-Item -Path HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail -ItemType Directory -Force }
New-ItemProperty HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Mail -Name 'Send Pictures With Document' -Value 1 -PropertyType 4 -Force
if (!(Test-Path -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail)) { New-Item -Path HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail -ItemType Directory -Force }
New-ItemProperty HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\Mail -Name 'Send Pictures With Document' -Value 1 -PropertyType 4 -Force

# check Office 2016 is installed
If (Test-Path -Path'C:\Program Files (x86)\Microsoft Office') 
{
Write-host "Setting signature for Office 2016"-ForegroundColor Green

    If ($ForceSignature -eq '0')
    {
    Write-host "Setting signature for Office 2016 as available" -ForegroundColor Green

    $MSWord = New-Object -comobject word.application
    $EmailOptions = $MSWord.EmailOptions
    $EmailSignature = $EmailOptions.EmailSignature
    $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
    $EmailSignature.NewMessageSignature="$SignatureName"
    $EmailSignature.ReplyMessageSignature="$SignatureName"

    }

    If ($ForceSignature -eq '1')
    {
    Write-Host "Setting signature for Office 2016 "
        If (!(Get-ItemProperty -Name 'NewSignature' -Path 'C:\ProgramData\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue))  
        { 
        New-ItemProperty 'C:\ProgramData\Microsoft\Office\16.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
        } 

        If (!(Get-ItemProperty -Name 'ReplySignature' -Path 'C:\ProgramData\Microsoft\Office\16.0\Common\MailSettings' -ErrorAction SilentlyContinue))  
        { 
        New-ItemProperty 'C:\ProgramData\Microsoft\Office\16.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
        } 
    }
}
