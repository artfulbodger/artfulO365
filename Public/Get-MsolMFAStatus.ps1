#Requires -module Msonline
function Get-MsolMFAStatus
{
  <#
    .SYNOPSIS
      Gets the Strong Authentication data for a given AzureAD User.

      .DESCRIPTION
      Connects to MSOnline service if a valid connection isn't available and gets the Strong AUthention status and default method for the given AzureAD user.

      .PARAMETER userprincipalname
      User Principal Name for the AzureAD user you wish to query.

      .EXAMPLE
      Get-MSOLMFAStatus -userprincipalname john.smith@example.com
      Gets the Display Name, UserPrincipalName, MFA Status and Enrolled MFA Method for john.smith@example.com

      .NOTES
      Requires the MSOnline PowerShell module.
      Requires Company Administrator role in AzureAD to be able to query MFA status and method.
      https://docs.microsoft.com/en-us/powershell/module/msonline/get-msoluser
  #>

      [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)][string]$userprincipalname
    )

    Begin
    {
        Try{
          Get-MsolCompanyInformation -ErrorAction Stop | Out-Null
        } Catch {
          Connect-msolservice -ErrorAction Stop
        }
    }
    Process
    {
      $msoluser = Get-Msoluser -UserPrincipalName $userprincipalname
      If($null -ne $msoluser.StrongAuthenticationRequirements.State){
        $method = ($msoluser.StrongAuthenticationMethods | Where-Object {$_.IsDefault -eq $true} | Select-Object MethodType).MethodType
      } else {
        $method = 'Not Enrolled'
      }
      $mfaresult += [pscustomobject]@{'Name' = $($msoluser.DisplayName)
                        'UserPrincipalName' = $($msoluser.UserPrincipalName)
                        'MFA'=If($null -eq $msoluser.StrongAuthenticationRequirements.State){'Disabled'}else{$msoluser.StrongAuthenticationRequirements.State}
                        'MFA Method' = $method}
    }
    End
    {
      $mfaresult
    }
}