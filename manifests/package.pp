# Class msoffice::package
#
# This class installs the Microsoft Office on windows
#
# Parameters:
#   [*ensure*]          - Control the existence of office    
#   [*deployment_root*] - The network location to go and find the package
#   [*version*]         - The version of office to install
#   [*edition*]         - The edition of office
#   [*arch*]            - The architecture version of office
#   [*products*]        - The list of products to install as part of the office suite
#   [*license_key*]     - The license key required to install
#
# Actions:
#
# Requires:
#
# Usage:
#
class msoffice::package(
  $ensure = 'present',
  $deployment_root = hiera('windows_deployment_root'),
  $version = hiera('msoffice_version'), 
  $edition = hiera('msoffice_edition'),
  $arch = 'x86', 
  $products = [],
  $license_key = ''
) {

  $temp_dir = 'C:\\Windows\\Temp'
  
  case $edition {
    'Standard': {
      if $version == '2007' {
        $office_product = 'Standard'
      } else {
        $office_product = 'Standardr'
      }
    }
    'Professional': {
      $office_product = 'Pro'
    }
    'Professional Plus': {
      if $version == '2003' {
        notify { "office ${version} does not have an edition: ${edition}": }
      }
      elsif $version == '2007' {
        $office_product = 'ProPlus'
      } else {
        $office_product = 'ProPlusr'
      }
    }
    'Enterprise': {
      if $version == '2007' {
        $office_product = 'Enterprise'
      } else {
        notify { "office ${version} does not have an edition: ${edition}": }
      }
    }
    'Ultimate': {
      if $version == '2007' {
        $office_product = 'Ultimater'
      } else {
        notify { "office ${version} does not have an edition: ${edition}": }
      }
    }
    default: {
      notify { "office ${version} does not have an edition: ${edition}": }
    }
  }
    
  case $version {
    '2003': {
      $office_root = "${deployment_root}\\OFFICE11\\${edition}"
      $office_reg_key = 'HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90110409-6000-11D3-8CFE-0150048383C9}'
    }
    '2007': {
      $office_root = "${deployment_root}\\OFFICE12\\${edition}"
      $office_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90120000-002A-0000-1000-0000000FF1CE}'
    }
    '2010': {
      $office_root = "${deployment_root}\\OFFICE14\\${edition}\\${arch}"
      $office_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{91140000-0011-0000-1000-0000000FF1CE}'
    }
    '2013': {
      $office_root = "${deployment_root}\\OFFICE15\\${edition}\\${arch}"
    }
    default: {
      notify { "office does not have a version: ${version}": }
    }   
  }
  
  if $ensure == 'present' {
    if $version == '2003' {
      if $office_root {
        file { "${temp_dir}\\office_config.ini":
          content => template('msoffice/setup.ini.erb'),
          mode    => '0755',
          owner   => 'Administrator',
          group   => 'Administrators',
        }
                
        exec { 'install-office':
          command   => "${office_root}\\SETUP.EXE /settings \"${temp_dir}\\office_config.ini\"",
          provider  => windows,
          logoutput => true,
          subscribe => File["${temp_dir}\\office_config.ini"],
          require   => File["${temp_dir}\\office_config.ini"]
        }
      }
    } else {
      if $office_root {
        file { "${temp_dir}\\office_config.xml":
          content => template('msoffice/config.erb'),
          mode    => '0755',
          owner   => 'Administrator',
          group   => 'Administrators',
        }
                
        exec { 'install-office':
          command   => "\"${office_root}\\setup.exe\" /config \"${temp_dir}\\office_config.xml\"",
          provider  => windows,
          logoutput => true,
          creates   => 'C:\\Program Files\\Microsoft Office',
          require   => File["${temp_dir}\\office_config.xml"]
        }
        
        exec { 'upgrade-office':
          command   => "\"${office_root}\\setup.exe\" /modify ${office_product} /config \"${temp_dir}\\office_config.xml\"",
          provider  => windows,
          logoutput => true,
          subscribe => File["${temp_dir}\\office_config.xml"],
          require   => File["${temp_dir}\\office_config.xml"]
        }
      }
    }
  } elsif $ensure == 'absent' {
    if $version == '2003' {
      exec { 'uninstall-office':
        command   => "& ${office_root}\\setup.exe /x pro11.msi /qb",
        provider  => powershell,
        logoutput => true,
        onlyif    => "if (Get-Item -LiteralPath \'\\${office_reg_key}\' -ErrorAction SilentlyContinue).GetValue(\'DisplayVersion\')) { exit 1 }"
      }
            
      file { ["${temp_dir}\\office_config.ini","${temp_dir}\\office_install.log"]:
        ensure  => absent,
        require => Exec['uninstall-office']
      }
    } else {
      exec { 'uninstall-office':
        command   => "& ${office_root}\\setup.exe /uninstall ${office_product} /config \"${temp_dir}\\office_config.xml\"",
        provider  => powershell,
        logoutput => true,
        onlyif    => "if (Get-Item -LiteralPath \'\\${office_reg_key}\' -ErrorAction SilentlyContinue).GetValue(\'DisplayVersion\')) { exit 1 }"
      }
    }
  } else {
    notify { "do not understand ensure agrument: ${ensure}": }
  }
}