#Microsoft Office module for Puppet


##Overview

Puppet module to manage Microsoft Office on Windows (2003-2013)

This module is also available on the [Puppet Forge](https://forge.puppetlabs.com/liamjbennett/msoffice)

[![Build
Status](https://secure.travis-ci.org/liamjbennett/puppet-msoffice.png)](http://travis-ci.org/liamjbennett/puppet-msoffice)
[![Dependency
Status](https://gemnasium.com/liamjbennett/puppet-msoffice.png)](http://gemnasium.com/liamjbennett/puppet-msoffice)

##Module Description

The purpose of this module is to install the Microsoft Office suite and configure it's many service packs, tools, utilities and registry options.

##Setup

####How Office is installed

Installation will be performed from files stored on a pre-defined network shared in a known structure. This is described in more detail [Here]()
It will install the core office suite in addition to the service pack and a default language interface pack.
Language Interfact Packs (LIPs) and Service Packs (SPs) can only be installed and the whole office installation will need to be removed to change that configuration

####Setup Requirements
Office version 2003-2013 on all versions of Windows are supported

##Usage
First please read the [Wiki](https://github.com/liamjbennett/puppet-msoffice/wiki) regarding how we assume your network
share should be configured. Then installing office is as simple as:

    msoffice { 'office 2010':
      version     => '2010',
      edition     => 'Professional Pro',
      sp          => '1'
      license_key => 'XXX-XXX-XXX-XXX-XXX',
      products    => ['Word,'Excel]
      ensure      => present,
    }


##Reference
Some basic information, for more read the [Wiki](https://github.com/liamjbennett/puppet-msoffice/wiki)

###Defintions:

msoffice::package     - the core office suite installation <br/>
msoffice::servicepack - the service pack update for office <br/>
msoffice::lip         - the language interface pack for office <br/>


##Limitations
Support for only RTM versions <br/>
Support for only Retail/Volume editions <br/>


##Development
Copyright (C) 2013 Liam Bennett - <liamjbennett@gmail.com> <br/>
Distributed under the terms of the Apache 2 license - see LICENSE file for details. <br/>
Further contributions and testing reports are extremely welcome - please submit a pull request or issue on [GitHub](https://github.com/liamjbennett/puppet-msoffice) <br/>
Please read the [Wiki](https://github.com/liamjbennett/puppet-msoffice/wiki) as there is a lot of useful information and links that will help you understand this module <br/>

###TODO
* More fine grained feature installation/un-installation
* Office products not part of the default suite (Viso, Lync etc)
* Office viewers
* Compatibility packs
* Manage desktop and quick launch toolbar shortcuts
* Allow easy tweaking of office registry settings
* Support older (pre 2003) versions of Office - just for fun!

##Release Notes

__0.0.2__ <br/>
Added lots of testing and support for installing LIPs.

__0.0.1__ <br/>
The initial proof-of-concept version
