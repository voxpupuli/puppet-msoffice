# Author::    Song An BUI (mailto:songan.bui@gmail.com)
# Copyright:: Copyright (c) 2015 Song An BUI
# License::   MIT

# == Define msoffice::settings::common
#
# This definition configures common settings for MS Office Suite products
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
# [*disableAllActiveX*]
# Enable or disable ActiveX Controls
# Values: 0;1
#
# [*ufiControls*]
# UFI Controls for ActiveX components
# Values: 1;2;3;4;5;6
# Please refer to: https://support.microsoft.com/en-us/kb/827742
# or https://technet.microsoft.com/en-us/library/cc178946(v=office.12).aspx
#
define msoffice::settings::common (
  $version,
  $arch = 'x86',
  $disableAllActiveX = '0',
  $ufiControls = '1'
){
  include msoffice::params
  $office_num = $msoffice::params::office_versions[$version]['version']
  
  if ($::architecture=='x64' and $arch=='x86') {
    $office_reg_key = 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office'
  }
  else {
    $office_reg_key = 'HKLM:\SOFTWARE\Microsoft\Office'
  }

  registry::value { 'DisableAllActiveX':
    key  => "${office_reg_key}\\Common\\Security",
    type => 'dword',
    data => $disableAllActiveX,
  }

  registry::value { 'UFIControls':
    key  => "${office_reg_key}\\Common\\Security",
    type => 'dword',
    data => $ufiControls,
  }

}