# Author::    Liam Bennett (mailto:liamjbennett@gmail.com)
# Copyright:: Copyright (c) 2014 Liam Bennett
# License::   MIT

# == Define msoffice
#
# Module to manage the installation and configuration of Microsoft Office
#
# === Requirements/Dependencies
#
# Currently reequires the puppetlabs/stdlib module on the Puppet Forge in
# order to validate much of the the provided configuration.
#
# === Parameters
#
# [*version*]
# The version of office to install
#
# [*edition*]
# The edition of office to install
#
# [*sp*]
# The service pack update to apply
#
# [*license_key*]
# The license key required to install
#
# [*arch*]
# The architecture version of office
#
# [*products*]
# The list of products to install as part of the office suite
#
# [*lang_code]
# The language code of the default install language
#
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
#     ensure      => present,
#   }
#
define msoffice(
  $version,
  $edition,
  $sp,
  $license_key,
  $deployment_root,
  $arch = 'x86',
  $products = [],
  $lang_code = 'en-us',
  $ensure = 'present'
) {

  include msoffice::params

  validate_re($version,'^(2003|2007|2010|2013)$', 'The version agrument specified does not match a valid version of office')

  $edition_regex = join(keys($msoffice::params::office_versions[$version]['editions']), '|')
  validate_re($edition,"^${edition_regex}$", 'The edition argument does not match a valid edition for the specified version of office')

  validate_re($sp,'^[0-3]$', 'The service pack specified does not match 0-3')
  validate_re($license_key,'^[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}-[a-zA-Z0-9]{5}$', 'The license_key argument speicifed is not correctly formatted')
  validate_re($arch,'^(x86|x64)$', 'The arch argument specified does not match x86 or x64')
  validate_array($products)

  $lang_regex = join(keys($msoffice::params::lcid_strings), '|')
  validate_re($lang_code,"^${lang_regex}$", 'The lang_code argument does not specifiy a valid language identifier')

  validate_re($ensure,'^(present|absent)$', 'The ensure argument does not match present or absent')

  msoffice::package { "microsoft office ${version}":
    ensure          => $ensure,
    version         => $version,
    edition         => $edition,
    license_key     => $license_key,
    arch            => $arch,
    lang_code       => $lang_code,
    products        => $products,
    sp              => $sp,
    deployment_root => $deployment_root
  }

  if $ensure == 'present' {
    msoffice::servicepack { "microsoft office ${version} servicepack ${sp}":
      version         => $version,
      sp              => $sp,
      arch            => $arch,
      deployment_root => $deployment_root
    }

    msoffice::lip { "microsoft lip ${lang_code}":
      version   => $version,
      lang_code => $lang_code,
      arch      => $arch
    }
  }

}
