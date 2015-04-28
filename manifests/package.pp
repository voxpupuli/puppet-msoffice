# Author::    Liam Bennett (mailto:liamjbennett@gmail.com)
# Copyright:: Copyright (c) 2014 Liam Bennett
# License::   MIT

# == Define msoffice::package
#
# This definition installs Microsoft Office package on windows
#
# === Parameters
#
# [*version*]
# The version of office to install
#
# [*edition*]
# The edition of office
#
# [*license_key*]
# The license key required to install
#
# [*arch*]
# The architecture version of office
#
# [*lang_code*]
# The language code of the default install language
#
# [*products*]
# The list of products to install as part of the office suite

# [*sp*]
# The service pack version that will be installed

# [*ensure*]
# Ensure the existence of the office installation
#
# [*deployment_root*]
# The network location where the office installation media is stored
#
# === Examples
#
#  To install Word and Excel packages from Office 2010 SP1:
#
#   msoffice { 'office 2010':
#     version     => '2010',
#     edition     => 'Professional Pro',
#     sp          => '1'
#     license_key => 'XXX-XXX-XXX-XXX-XXX',
#     products    => ['Word,'Excel]
#     ensure      => present
#   }
#
define msoffice::package(
  $version,
  $edition,
  $license_key,
  $arch = 'x86',
  $lang_code = 'en-us',
  $products = ['Word','Excel','Powerpoint','Outlook'],
  $sp = '0',
  $ensure = 'present',
  $deployment_root = '',
  $company_name = '',
  $user_name = '',
) {

  include msoffice::params

  validate_re($version,'^(2003|2007|2010|2013)$', 'The version agrument specified does not match a valid version of office')

  $edition_regex = join(keys($msoffice::params::office_versions[$version]['editions']), '|')
  validate_re($edition,"^${edition_regex}$", 'The edition argument does not match a valid edition for the specified version of office')

  validate_re($license_key,'^[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}$', 'The license_key argument speicifed is not correctly formatted')
  validate_re($arch,'^(x86|x64)$', 'The arch argument specified does not match x86 or x64')

  $lang_regex = join(keys($msoffice::params::lcid_strings), '|')
  validate_re($lang_code,"^${lang_regex}$", 'The lang_code argument does not specifiy a valid language identifier')

  validate_re($ensure,'^(present|absent)$', 'The ensure argument does not match present or absent')
  validate_array($products)
  validate_re($sp,'^[0-3]$', 'The service pack specified does not match 0-3')

  $office_num = $msoffice::params::office_versions[$version]['version']
  $office_product = $msoffice::params::office_versions[$version]['editions'][$edition]['office_product']
  $product_key = $msoffice::params::office_versions[$version]['prod_key']
  
  if ($::architecture=='x64' and $arch=='x86') {
    $office_reg_key = "HKLM:\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\${office_num}.0\\Common\\InstallRoot"
  } 
  else {
    $office_reg_key = "HKLM:\\SOFTWARE\\Microsoft\\Office\\${office_num}.0\\Common\\InstallRoot"
  }
  $office_build = $msoffice::params::office_versions[$version]['service_packs'][$sp]['build']
  $office_hasMultipleArch = $msoffice::params::office_versions[$version]['hasMultipleArch']

  if $office_hasMultipleArch {
    $office_root = "${deployment_root}\\OFFICE${office_num}\\${edition}\\${arch}"
  } else {
    $office_root = "${deployment_root}\\OFFICE${office_num}\\${edition}"
  }

  if $office_root {
    file { "${msoffice::params::temp_dir}\\office${office_num}":
      ensure  => directory,
      source  => $office_root,
      recurse => true,
      purge   => true,
      mode    => '0777',
      owner   => 'Administrator',
      group   => 'Administrators',
    }
  }

  if $ensure == 'present' {
    if $version == '2003' {
      file { "${msoffice::params::temp_dir}\\office${office_num}_config.ini":
        content => template('msoffice/setup.ini.erb'),
        mode    => '0755',
        owner   => 'Administrator',
        group   => 'Administrators',
      }

      exec { "install-office${office_num}":
        path      => $::path,
        command   => "& cmd.exe /c start /w \"${msoffice::params::temp_dir}\\office${office_num}\\setup.exe\" /settings \"${msoffice::params::temp_dir}\\office${office_num}_config.ini\"",
        provider  => powershell,
        logoutput => true,
        timeout   => 0,
        unless    => template('msoffice/check_office_installed.ps1.erb'),
        subscribe => File["${msoffice::params::temp_dir}\\office${office_num}_config.ini"],
        require   => [File["${msoffice::params::temp_dir}\\office${office_num}_config.ini"],
                      File["${msoffice::params::temp_dir}\\office${office_num}"]],
      }
    } else {
      file { "${msoffice::params::temp_dir}\\office${office_num}_config.xml":
        content => template('msoffice/config.erb'),
        mode    => '0755',
        owner   => 'Administrator',
        group   => 'Administrators',
      }

      exec { "install-office${office_num}":
        path      => $::path,
        command   => "& cmd.exe /c start /w \"${msoffice::params::temp_dir}\\office${office_num}\\setup.exe\" /config \"${msoffice::params::temp_dir}\\office${office_num}_config.xml\"",
        provider  => powershell,
        logoutput => true,
        timeout   => 0,
        unless    => template('msoffice/check_office_installed.ps1.erb'),
        require   => [File["${msoffice::params::temp_dir}\\office${office_num}_config.xml"],
                      File["${msoffice::params::temp_dir}\\office${office_num}"]],
      }

      # exec { 'upgrade-office':
      #   command   => "& \"${office_root}\\office${office_num}\\setup.exe\" /modify ${office_product} /config \"${msoffice::params::temp_dir}\\office_config.xml\"",
      #   provider  => powershell,
      #   logoutput => true,
      #   subscribe => File["${msoffice::params::temp_dir}\\office_config.xml"],
      #   require   => [File["${msoffice::params::temp_dir}\\office_config.xml"],
      #                 File["${msoffice::params::temp_dir}\\office"]],
      # }
    }
  } elsif $ensure == 'absent' {
    if $version == '2003' {
      exec { "uninstall-office${office_num}":
        path      => $::path,
        command   => "& cmd.exe /c start /w \"${msoffice::params::temp_dir}\\office${office_num}\\setup.exe\" /x ${office_product}.msi /qb",
        provider  => powershell,
        logoutput => true,
        timeout   => 0,
        onlyif    => template('msoffice/check_office_installed.ps1.erb'),
        require   => File["${msoffice::params::temp_dir}\\office${office_num}"],
      }

      file { ["${msoffice::params::temp_dir}\\office${office_num}_config.ini","${msoffice::params::temp_dir}\\office_install.log"]:
        ensure  => absent,
        require => Exec["uninstall-office${office_num}"]
      }
    } else {
      exec { "uninstall-office${office_num}":
        path      => $::path,
        command   => "& cmd.exe /c start /w \"${msoffice::params::temp_dir}\\office${office_num}\\setup.exe\" /uninstall ${office_product} /config \"${msoffice::params::temp_dir}\\office${office_num}_config.xml\"",
        provider  => powershell,
        logoutput => true,
        timeout   => 0,
        onlyif    => template('msoffice/check_office_installed.ps1.erb'),
        require   => File["${msoffice::params::temp_dir}\\office${office_num}"],
      }
    }
  } else { }
}
