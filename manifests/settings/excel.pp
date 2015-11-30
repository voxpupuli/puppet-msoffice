# Author::    Song An BUI (mailto:songan.bui@gmail.com)
# Copyright:: Copyright (c) 2015 Song An BUI
# License::   MIT

# == Define msoffice::settings::excel
#
# This definition configures settings for MS Office Excel
#
# === Parameters
#
# [*version*]
# The version of office to configure
#
# [*arch*]
# The architecture of MS Office Suite
# Values: x86;x64
#
# [*VBAWarnings*]
# Set security level for VBA Warning messages
# Values: 1;2;3;4 
# Please refer to: https://technet.microsoft.com/en-us/library/cc178946(v=office.12).aspx
#
# [*AccessVBOM*]
# Enable or disable access to VB Object Model
# Values: 0;1
#
define msoffice::settings::excel (
  $version,
  $arch = 'x86',
  $vbaWarnings = '1',
  $accessVBOM = '1'
){
  include msoffice::params
  $office_num = $msoffice::params::office_versions[$version]['version']
  
  if ($::architecture=='x64' and $arch=='x86') {
    $office_reg_key = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office'
    $office_reg_hkcu_key = 'HKCU:\SOFTWARE\Wow6432Node\Microsoft\Office'
  }
  else {
    $office_reg_key = 'HKLM:\SOFTWARE\Microsoft\Office'
    $office_reg_hkcu_key = 'HKCU:\SOFTWARE\Microsoft\Office'
  }

  registry::value { 'VBAWarnings':
    key  => "${office_reg_key}\\${office_num}.0\\Excel\\Security",
    type => 'dword',
    data => $vbaWarnings,
  }
  # VBAWarnings does not seem to work system-wide. Must use HKCU.
  # $excel_vbawarnings_key = "${office_reg_hkcu_key}\\${office_num}.0\\Excel\\Security"
  # $excel_vbawarnings_name = 'VBAWarnings'
  # $excel_vbawarnings_value = $vbaWarnings
  # $excel_vbawarnings_scriptpath = "${msoffice::params::temp_dir}\\office${office_num}_set_excel_vbawarnings.ps1"
  # # exec { 'Excel VBAWarnings HKCU':
  # #   path     => $::path,
  # #   provider => powershell,
  # #   command  => template('msoffice/set_excel_vbawarnings.ps1.erb'),
  # #   unless   => template('msoffice/check_excel_vbawarnings.ps1.erb'),
  # # }
  # file { 'set excel vba warnings script':
  #   ensure             => present,
  #   path               => $excel_vbawarnings_scriptpath,
  #   source_permissions => ignore,
  #   content            => template('msoffice/set_excel_vbawarnings.ps1.erb'),
  # }
  # ->
  # exec { 'set excel vba warnings':
  #   path        => $::path,
  #   command     => "${devxexec::path} /user:${devxexec::username} /password:${devxexec::password} /sessionid:1 \"powershell -ExecutionPolicy unrestricted ${iescript_path}\"",
  #   refreshonly => true,
  # }

  registry::value { 'AccessVBOM':
    key  => "${office_reg_key}\\${office_num}.0\\Excel\\Security",
    type => 'dword',
    data => $accessVBOM,
  }

}