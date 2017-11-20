# MS Office module for Puppet

[![Build Status](https://travis-ci.org/voxpupuli/puppet-msoffice.png?branch=master)](https://travis-ci.org/voxpupuli/puppet-msoffice)
[![Code Coverage](https://coveralls.io/repos/github/voxpupuli/puppet-msoffice/badge.svg?branch=master)](https://coveralls.io/github/voxpupuli/puppet-msoffice)
[![Puppet Forge](https://img.shields.io/puppetforge/v/puppet/msoffice.svg)](https://forge.puppetlabs.com/puppet/msoffice)
[![Puppet Forge - downloads](https://img.shields.io/puppetforge/dt/puppet/msoffice.svg)](https://forge.puppetlabs.com/puppet/msoffice)
[![Puppet Forge - endorsement](https://img.shields.io/puppetforge/e/puppet/msoffice.svg)](https://forge.puppetlabs.com/puppet/msoffice)
[![Puppet Forge - scores](https://img.shields.io/puppetforge/f/puppet/msoffice.svg)](https://forge.puppetlabs.com/puppet/msoffice)

#### Table of Contents

1. [Overview](#overview)
2. [Module Description - What the module does and why it is useful](#module-description)
3. [Setup - The basics of getting started with msoffice](#setup)
    * [What msoffice affects](#what-msoffice-affects)
    * [Beginning with msoffice](#beginning-with-msoffice)
4. [Usage - Configuration options and additional functionality](#usage)
5. [Reference - An under-the-hood peek at what the module is doing and how](#reference)
5. [Limitations - OS compatibility, etc.](#limitations)
6. [Development - Guide for contributing to the module](#development)

## Overview

Puppet module to manage Microsoft Office on Windows (2003-2016)

## Module Description

The purpose of this module is to install the Microsoft Office suite and configure
it's many service packs, tools, utilities and registry options.

## Setup

### What msoffice affects

* Installs packages for each office product
* Installs package for the Service Pack (if configured)
* Installs packages for each language pack (if configured)

### Beginning with msoffice

  To install Word and Excel packages from Office 2010 SP1:

```puppet
    msoffice { 'office 2010':
      version     => '2010',
      edition     => 'Professional Pro',
      sp          => '1',
      license_key => 'XXX-XXX-XXX-XXX-XXX',
      products    => ['Word,'Excel'],
      ensure      => present
    }
```

## Usage

### Classes and Defined Types

#### Defined Type: `msoffice`

The primary definition of the msoffice module. It will install office products,
language packs and updates.

**Parameters within `msoffice`:**

##### `version`

The version of office to install

##### `edition`

The edition of office to install

##### `sp`

The service pack update to apply

##### `license_key`

The license key required to install

##### `arch`

The architecture version of office

##### `products`

The list of products to install as part of the office suite

##### `lang_code`

The language code of the default install language

##### `ensure`

Ensure the existence of the office installation

##### `deployment_root`

The network location where the office installation media is stored

#### Defined Type: `msoffice::package`

The definition which installs the main office products.

**Parameters within `msoffice::package`:**

##### `version`

The version of office to install

##### `edition`

The edition of office to install

##### `license_key`

The license key required to install

##### `arch`

The architecture version of office

##### `lang_code`

The language code of the default install language

##### `products`

The list of products to install as part of the office suite

##### `sp`

The service pack update to apply

##### `ensure`

Ensure the existence of the office installation

##### `deployment_root`

The network location where the office installation media is stored

#### Defined Type: `msoffice::lip`

The definition which installs language interface packs into an existing office installation

**Parameters within `msoffice::lip`:**

##### `version`

The version of office that was installed

##### `lang_code`

The language code of the language to install

##### `arch`

The architecture version of office

##### `deployment_root`

The network location where the office installation media is stored

#### Defined Type: `msoffice::servicepack`

The definition which installs service packs into an existing office installation

**Parameters within `msoffice::servicepack`:**

##### `version`

The version of office

##### `sp`

The service pack update to install

##### `arch`

The architecture version of office

##### `deployment_root`

The network location where the office installation media is stored

## Reference

### Defined Types

#### Public Defined Types

* [`msoffice`](#define_package): The core office suite installation
* [`msoffice::package`](#define_package): The core office suite installation
* [`msoffice::servicepack`](#define_servicepack): The service pack update for office
* [`msoffice::lip`](#define_lip): The language interface pack for office

## Limitations

This module is tested on the following platforms:

* Windows 2008 R2

It is tested with the OSS version of Puppet only.

Support for only RTM versions
Support for only Retail/Volume editions

## Development

### Contributing

Please read CONTRIBUTING.md for full details on contributing to this project.
