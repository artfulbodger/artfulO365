#Requires -Version 5.1
#Requires -Modules @{ ModuleName="Microsoft.Online.SharePoint.PowerShell"; ModuleVersion="16.0.8924.0" }
#Requires -Modules ActiveDirectory 

function Request-OneDriveProvisioning
{
  <#
      .SYNOPSIS
      Requests One Drive Storage it if has not been provisioned.

      .DESCRIPTION
      Checks if OneDrive storage has been provisioned - if not the storage is requested.

      .EXAMPLE
      Request-OneDriveProvisioning -orgname contoso -Server dc1.contoso.com -OrganizationalUnit 'OU=Users,DC=contoso,DC=com'
      Requests OneDrive storage provisioning for all users in the 'Users' OU where the SPOSite does not exist.

      .NOTES
      By default, the first time that a user browses to their OneDrive it's automatically provisioned for them. In some cases you might want your users' OneDrive locations to be pre-provisioned.

      .LINK
      https://docs.microsoft.com/en-us/onedrive/pre-provision-accounts
  #>

      [CmdletBinding(DefaultParameterSetName='OrganizationalUnit')]
    Param
    (
        [Parameter(Mandatory,HelpMessage='Office365 Tenant name')][string]$orgname,
        [Parameter(Mandatory,HelpMessage='Domain Controller to use')][string]$Server,
        [Parameter(Mandatory,HelpMessage='DN of OU to process',ParameterSetName='OrganizationalUnit')][ValidateScript({
          Try {
            Get-ADOrganizationalUnit -Identity $_ -ErrorAction Stop | out-null
            return $true
          } Catch {
            throw "Organizational Unit could not be found"
          }
        })][string]$OrganizationalUnit
    )

    Begin
    {
      Import-Module -Name Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
      Connect-SPOService -Url https://$orgName-admin.sharepoint.com
    }
    Process
    {
      Switch($PSCmdlet.ParameterSetName){
        'OrganizationalUnit'{$od4b = Get-Aduser -server $Server -LDAPFilter '(&(objectclass=user)(objectcategory=person))' -SearchBase $OrganizationalUnit -ResultPageSize 100 -Properties mail}
      }  
      Foreach($user in $od4b){
        Try{
          $user_uri = ($($user.mail).replace("@","_")).replace(".","_")
          $response = Invoke-webrequest -Uri https://$orgName-my.sharepoint.com/personal/$user_uri -ErrorAction Stop
        } catch [Net.WebException] {
          $response = $_.Exception.Response
        }
        If ($Response.StatusCode -eq 200 -or $Response.StatusCode -eq 403 )   { 
          Write-Host -ForegroundColor Green "Site for user $($user.mail) exists!"
        }
        Else {
          Write-Host -ForegroundColor Yellow "Site for user $($user.mail) does not respond, attempting to provision"
          Request-SPOPersonalSite -UserEmails $user.mail
        }
    
      }
    }
    End
    {
    }
}
