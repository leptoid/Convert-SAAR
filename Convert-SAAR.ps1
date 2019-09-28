<#
.SYNOPSIS
    Parse DD Form 2875 for user account creation data, using iTextSharp library.
.DESCRIPTION
    Parse DD Form 2875 for user account creation data, using iTextSharp library.
.NOTES
    Author: JBear
    Date: 11/10/2018
    Version: 1.0
    Notes: Requires iTextSharp library.
    
    Modified by Nathan Kramer
    Date 09/28/2019
    Version 1.1
#>

#Load iTextSharp .NET library
try {
    
    #Set local path to itextsharp.dll; *Hint: Can be saved anywhere you want/not restricted to System32.
    #Allowing the computer to do the work; comment out next 2 line if script acts erratic
    $itextsharp = ls -Recurse .\itextsharp.dll
    Add-Type -Path $itextsharp[0]
    #Uncomment below to set manually (when auto is acting erratic/taking too long
    #Add-Type -Path 'D:\Scripts\tabula\.powershell\test\itextsharp.dll'
}

catch {

   Write-Error $_
   Break
}

#Retrieve files from specified SAAR directory
$Files = (Get-ChildItem 'D:\Scripts\tabula\.powershell\test\Tickets\').FullName

foreach($File in $Files) {

    #Open PDF object
    $PDF = New-Object iTextSharp.text.pdf.PdfReader -ArgumentList $File

    #Retrieve required fields and data
    $Data = $PDF.AcroFields.XFA.DatasetsNode.Data.topmostSubform

    #Split name into first, middle, and last
    $Name = $Data.Name.Split(',').Split(" ") | where {$_ -ne ''}


    #Verify user signature
    $UserSignature = $PDF.AcroFields.VerifySignature("usersign")

    [PSCUstomObject] @{
    
        #User first name
        Firstname = $Name[1].Trim()

        #User last name
        LastName = $Name[0].Trim()

        #User middle initial
        MiddleIn =            
        #Error handling for null middle initials
        if(!([String]::IsNullOrWhiteSpace($Name[2]))) {
          
            $Name[2].Trim()
        }

        else {
            $Name[2] 
        }
        
        #User CAC/EDI number
        UserID = 
        
        if(!([String]::IsNullOrWhiteSpace($UserSignature.SignName))) {
            
            $UserSignature.SignName.Split(".") | Select -Last 1
        }

        #If SAAR form is not signed by user
        else {
            
            $Data.UserID.Replace('EDIPI','').Replace(' ','').Replace('#','').Replace(':','').Trim()
        }

        #User email address
        Email = $Data.ReqEmail.Trim()

        #User phone number
        Phone = $Data.ReqPhone.Trim()

        #User organization
        Organization = $Data.ReqOrg.Trim()

        #User job title
        JobTitle = $Data.ReqTitle.Trim()
    }
    
    $PDF.Close()
}
