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
# Enable or disable VBA Warning messages
# Values: 0;1
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
  }
  else {
    $office_reg_key = 'HKLM:\SOFTWARE\Microsoft\Office'
  }

  registry::value { 'VBAWarnings':
    key  => "${office_reg_key}\\${office_num}.0\\Excel\\Security",
    type => 'dword',
    data => $vbaWarnings,
  }

  registry::value { 'AccessVBOM':
    key  => "${office_reg_key}\\${office_num}.0\\Excel\\Security",
    type => 'dword',
    data => $accessVBOM,
  }

}