# Class msoffice::servicepack
#
# This class installs the Microsoft Office service pack update
#
# Parameters:
#   [*ensure*]          - Control the existence of the service pack  
#   [*deployment_root*] - The network location to go and find the package
#   [*version*]         - The version of office
#   [*arch*]            - The architecture version of office
#   [*sp*]              - The service pack update to install
#
# Actions:
#
# Requires:
#
# Usage:
#
class msoffice::servicepack(
  $ensure = 'present',
  $deployment_root = hiera('windows_deployment_root'),
  $version = hiera('msoffice_version'), 
  $sp = ''
) {
  
  $sp_install_args = '/q /norestart'
  
  case $version {
    '2003': {
      case $sp {
        '1': { 
          $setup = 'Office2003SP1-kb842532-fullfile-enu'
          if $::architecture == 'x64' {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90120000-0018-0409-0000-0000000FF1CE}_PROPLUS_{4CA4ECC1-DBD4-4591-8F4C-AA12AD2D3E59}'
          }
          else {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90120000-0018-0409-0000-0000000FF1CE}_PROPLUS_{4CA4ECC1-DBD4-4591-8F4C-AA12AD2D3E59}'
          }
        }
        '2': { 
          $setup = 'Office2003SP2-KB887616-FullFile-ENU' 
          if $::architecture == 'x64' {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\{90120000-0011-0000-0000-0000000FF1CE}_PROPLUS_{0B36C6D6-F5D8-4EAF-BF94-4376A230AD5B}'
          }
          else {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90120000-0011-0000-0000-0000000FF1CE}_PROPLUS_{0B36C6D6-F5D8-4EAF-BF94-4376A230AD5B}'
          }
        }
        '3': { 
          $setup = 'Office2003SP3-KB923618-FullFile-ENU'
          if $::architecture == 'x64' {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90120000-0011-0000-0000-0000000FF1CE}_PROPLUS_{6E107EB7-8B55-48BF-ACCB-199F86A2CD93}'
          }
          else {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90120000-0011-0000-0000-0000000FF1CE}_PROPLUS_{6E107EB7-8B55-48BF-ACCB-199F86A2CD93}'
          } 
        }
        default: { fail("office ${version} does have have service pack ${sp}") }
      }

      $install_args = '/q'
      $sp_root = "${deployment_root}\\OFFICE11\\SPs"
    }
    '2007': {
      case $sp {
        '1': { $setup = 'office2007sp1-kb936982-fullfile-en-us.exe' }
        '2': { $setup = 'office2007sp2-kb953195-fullfile-en-us.exe' }
        '3': { $setup = 'office2007sp3-kb2526086-fullfile-en-us.exe' }
        default: { fail("office ${version} does have have service pack ${sp}") }
      }
      
      $sp_root = "${deployment_root}\\OFFICE12\\SPs"
    }
    '2010': {
      case $sp {
        '1': { 
          $setup = "officesuite2010sp1-kb2460049-${::arch}-fullfile-en-us.exe"
          if ($::architecture == 'x64') {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90140000-0016-0409-0000-0000000FF1CE}_Office14.PROPLUSR_{6BD185A0-E67F-4F77-8BCD-E34EA6AE76DF}' 
          }
          else {
            $sp_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90140000-0016-0409-0000-0000000FF1CE}_Office14.PROPLUSR_{6BD185A0-E67F-4F77-8BCD-E34EA6AE76DF}'
          }
        }
        default: { fail("office ${version} does have have service pack ${sp}") }
      }
      $sp_root = "${deployment_root}\\OFFICE14\\SPs"
    }
    '2013': {
      $sp_root = "${deployment_root}\\OFFICE15\\SPs"
    }
    default: { fail("office does not have a version: ${version}") }
  }
  
  if $ensure == 'present' {
    exec { 'install-sp':
      command   => "& ${sp_root}\\${setup} ${sp_install_args}",
      provider  => powershell,
      logoutput => true,
      onlyif    => "if (Get-Item -LiteralPath \'\\${sp_reg_key}\' -ErrorAction SilentlyContinue).GetValue(\'DisplayVersion\')) { exit 1 }"
    }   
  } else {
    #TODO uninstall
  }
}