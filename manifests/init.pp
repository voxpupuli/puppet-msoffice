# Define msoffice
#
# This definition installs the Microsoft Office on windows with a specified service pack
#
# Parameters:
#   [*version*]         - The version of office to install
#   [*edition*]         - The edition of office
#   [*sp*]              - The service pack update to apply
#   [*license_key*]     - The license key required to install
#   [*arch*]            - The architecture version of office
#   [*products*]        - The list of products to install as part of the office suite
#   [*lang_code]        - The language code of the default install language
#   [*ensure*]          - Control the existence of office
#
# Actions:
#
# Requires:
#
# Usage:
#
#   msoffice { 'office 2010':
#     ensure      => present,
#     version     => '2010',
#     edition     => 'Professional Pro',
#     sp          => '1'
#     license_key => 'XXX-XXX-XXX-XXX-XXX'
#   }
define msoffice(
  $version,
  $edition,
  $sp,
  $license_key,
  $arch = 'x86',
  $products = [],
  $lang_code = 'en-us',
  $ensure = 'present',
) {

  include msoffice::params

  validate_re($version,'^(2003|2007|2010|2013)$', 'The version agrument specified does not match a valid version of office')

  $edition_regex = join(keys($msoffice::params::office_versions[$version]['editions']), "|")
  validate_re($edition,"^${edition_regex}$", 'The edition argument does not match a valid edition for the specified version of office')

  validate_re($sp,'^[0-3]$', 'The service pack specified does not match 0-3')
  validate_re($license_key,'^[a-zA-Z]{5}-[a-zA-Z]{5}-[a-zA-Z]{5}-[a-zA-Z]{5}-[a-zA-Z]{5}$', 'The license_key argument speicifed is not correctly formatted')
  validate_re($arch,'^(x86|x64)$', 'The arch argument specified does not match x86 or x64')
  validate_array($products)

  $lang_regex = join($msoffice::params::lcid_strings, "|")
  validate_re($lang_code,"^${lang_regex}$", 'The lang_code argument does not specifiy a valid language identifier')

  validate_re($ensure,'^(present|absent)$', 'The ensure argument does not match present or absent')

  msoffice::package { "microsoft office ${version}":
    version     => $version,
    edition     => $edition,
    license_key => $license_key,
    arch        => $arch,
    lang_code   => $lang_code,
    products    => $products,
    sp          => $sp,
    ensure      => $ensure,
  }

  if $ensure == 'present' {
    msoffice::servicepack { "microsoft office ${version} servicepack ${sp}":
      version     => $version,
      sp          => $sp,
      arch        => $arch,
    }
  }

}