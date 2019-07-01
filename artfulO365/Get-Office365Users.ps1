function Get-Office365Users
{
  <#
    .SYNOPSIS
    Describe purpose of "Get-Office365Users" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .EXAMPLE
    Get-Office365Users
    Describe what this call does

    .NOTES
    Place additional notes here.

    .LINK
    URLs to related sites
    The first link is opened by Get-Help -Online Get-Office365Users

    .INPUTS
    List of input types that are accepted by this function.

    .OUTPUTS
    List of output types produced by this function.
  #>


    [CmdletBinding()]

    Begin
    {
      Connect-AzureAD
      $tenantuserlist = @()
    }
    Process
    {
      $aduserlist = Get-AzureADUser -all $true | where-object {$_.usertype -eq 'Member'}
      Foreach ($aduser in $aduserlist){
        $tenantuserlist += [pscustomobject]@{'Firstname'=$aduser.GivenName
          'Surname'=$aduser.Surname
          'Display Name'=$aduser.DisplayName
          'Company Name'=$aduser.CompanyName
          'Department'=$aduser.Department
          'Title'=$aduser.JobTitle}
      }
      $tenantuserlist | Export-Csv -Path "$env:SystemDrive\Users\Artfulbodger\OneDrive - Devonshire County Scout Council\O365Users.csv" -NoClobber -NoTypeInformation
    }
    End
    {
    }
}