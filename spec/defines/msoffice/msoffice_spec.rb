require 'spec_helper'

describe 'msoffice', type: :define do
  lcid_strings = YAML.load_file(Dir.pwd + '/spec/lcid_strings.yml')

  describe 'installing with unknown version' do
    let :title do 'msoffice with unknown version' end
    let(:params) do
      {
        version: 'xxx',
        edition: 'Standard',
        sp: '1',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect do
        should contain_msoffice__package('microsoft office 2010')
      end.to raise_error
    end
  end

  %w(2003 2007 2010 2013).each do |version|
    describe "installing #{version} with wrong edition" do
      let :title do "msoffice for #{version}" end
      let(:params) do
        {
          version: version,
          edition: 'fubar',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it do
        expect do
          should contain_msoffice__package("microsoft office #{version}")
        end.to raise_error
      end
    end
  end

  describe 'installing with unknown sp' do
    let :title do 'msoffice unknown sp version' end
    let(:params) do
      {
        version: '2010',
        edition: 'Standard',
        sp: '5',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect do
        should contain_msoffice__package('microsoft office 2010')
      end.to raise_error
    end
  end

  describe 'incorrect license key' do
    let :title do 'msoffice with incorrect license key' end
    let(:params) do
      {
        version: '2010',
        edition: 'Standard',
        sp: '1',
        license_key: 'XXXXX-XXXXX-XXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect do
        should contain_msoffice__package('microsoft office 2010')
      end.to raise_error
    end
  end

  describe 'incorrect arch' do
    let :title do 'msoffice with incorrect arch' end
    let(:params) do
      {
        arch: 'fubar',
        version: '2010',
        edition: 'Standard',
        sp: '1',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect do
        should contain_msoffice__package('microsoft office 2010')
      end.to raise_error
    end
  end

  describe 'incorrect ensure' do
    let :title do 'msoffice with incorrect ensure' end
    let(:params) do
      {
        ensure: 'fubar',
        version: '2010',
        edition: 'Standard',
        sp: '1',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect do
        should contain_msoffice__package('microsoft office 2010')
      end.to raise_error
    end
  end

  lcid_strings.keys.each do |lang_code|
    describe "valid country: #{lang_code}" do
      let :title do 'msoffice with valid countries' end
      let(:params) do
        {
          lang_code: lang_code,
          version: '2010',
          edition: 'Standard',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it do
        expect do
          should contain_msoffice__package('microsoft office 2010')
        end.not_to raise_error
      end
    end
  end

  [2003, 2007, 2010].each do |version|
    # so we need version as a string
    # rubocop yells at us when we put strings in the array
    version = version.to_s
    describe "installs the core package for #{version}" do
      let :title do "office #{version}" end
      let(:params) do
        {
          version: version,
          edition: 'Standard',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it { should contain_msoffice__package("microsoft office #{version}") }
    end

    describe "installs the service pack when office #{version} present" do
      let :title do "office #{version}" end
      let(:params) do
        {
          ensure: 'present',
          version: version,
          edition: 'Standard',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it { should contain_msoffice__servicepack("microsoft office #{version} servicepack 1") }
    end

    describe "does not install the service pack when office #{version} absent" do
      let :title do "office #{version}" end
      let(:params) do
        {
          ensure: 'absent',
          version: version,
          edition: 'Standard',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it { should_not contain_msoffice__servicepack("microsoft office #{version} servicepack 1") }
    end

    describe "installs the lip when office #{version} present" do
      let :title do "office #{version}" end
      let(:params) do
        {
          lang_code: 'de-de',
          ensure: 'present',
          version: version,
          edition: 'Standard',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it { should contain_msoffice__lip('microsoft lip de-de') }
    end

    describe "does not install the lip when office #{version} present" do
      let :title do "office #{version}" end
      let(:params) do
        {
          lang_code: 'de-de',
          ensure: 'absent',
          version: version,
          edition: 'Standard',
          sp: '1',
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\test-server\packages'
        }
      end

      it { should_not contain_msoffice__lip('microsoft lip de-de') }
    end
  end
end
