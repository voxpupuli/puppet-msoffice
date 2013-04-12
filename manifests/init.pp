# Class msoffice
#
# This class installs the Microsoft Office on windows with a specified service pack
#
# Parameters:
#   [*ensure*]          - Control the existence of office
#   [*deployment_root*] - The network location to go and find the package
#   [*version*]         - The version of office to install
#   [*edition*]         - The edition of office
#   [*arch*]            - The architecture version of office
#   [*products*]        - The list of products to install as part of the office suite
#   [*license_key*]     - The license key required to install
#   [*sp*]              - The service pack update to apply
#
# Actions:
#
# Requires:
#
# Usage:
#
#   class { '_msoffice': 
#     ensure => present,
#     version => '2010',
#     edition => 'Professional Pro',
#     license_key => 'XXX-XXX-XXX-XXX-XXX'
#   }
class windows_msoffice(
  $ensure = 'present', 
  $deployment_root = hiera('windows_deployment_root'),
  $version = hiera('msoffice_version'), 
  $edition = hiera('msoffice_edition'),
  $arch = 'x86',
  $products = [],
  $license_key = hiera('msoffice_license_key'),
  $sp = hiera('msoffice_servicepack')
) {

  class { 'msoffice::package':
    ensure      => $ensure,
    version     => $version,
    edition     => $edition,
    arch        => $arch,
    products    => $products,
    license_key => $license_key
  }
  
  $sp_status = $sp ? {
    '0'     => absent,
    default => present
  }

  class { 'msoffice::servicepack':
    ensure  => $sp_status,
    version => $version,
    sp      => $sp
  }
    
}