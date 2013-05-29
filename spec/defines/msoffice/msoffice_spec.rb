require 'spec_helper'

describe 'msoffice', :type => :define do
  let :hiera_data do
    {
        :windows_deployment_root => '\\test-server\packages',
        :company_name => 'Example Inc',
        :office_username => 'Joe'
    }
  end

  ['2003','2007','2010'].each do |version|
    describe "installs the core package for #{version}" do
      let :title do "office #{version}" end
      let :params do
        { :version => version, :edition => 'Standard', :sp => '1', :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should contain_msoffice__package("microsoft office #{version}")}
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "installs the service pack when office #{version} present" do
      let :title do "office #{version}" end
      let :params do
        { :ensure => 'present', :version => version, :edition => 'Standard', :sp => '1', :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should contain_msoffice__servicepack("microsoft office #{version} servicepack 1")}
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "does not install the service pack when office #{version} absent" do
      let :title do "office #{version}" end
      let :params do
        { :ensure => 'absent', :version => version, :edition => 'Standard', :sp => '1', :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should_not contain_msoffice__servicepack("microsoft office #{version} servicepack 1")}
    end
  end

end