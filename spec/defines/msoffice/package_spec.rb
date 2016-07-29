require 'spec_helper'

office_versions = YAML.load_file(Dir.pwd + '/spec/office_versions.yml')
lcid_strings = YAML.load_file(Dir.pwd + '/spec/lcid_strings.yml')

describe 'msoffice::package', type: :define do
  office_versions['2013']['editions'].keys.each do |edition|
    product = office_versions['2013']['editions'][edition]['office_product']

    describe "installing office 2013 #{edition}" do
      let :title do 'msoffice for 2013' end
      let(:params) do
        {
          version: '2013',
          edition: edition,
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\\\test-server\\packages'
        }
      end

      it do
        should contain_exec('install-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE15\\#{edition}\\setup.exe\" /config \"C:\\Windows\\Temp\\office_config.xml\"",
          'provider' => 'windows'
      )
      end

      it do
        should contain_exec('upgrade-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE15\\#{edition}\\setup.exe\" /modify #{product} /config \"C:\\Windows\\Temp\\office_config.xml\""
      )
      end
    end

    office_versions['2013']['service_packs'].keys.each do |sp|
      build = office_versions['2013']['service_packs'][sp]['build']
      describe "uninstalling office 2013 #{edition} SP#{sp}" do
        let :title do 'msoffice for 2013' end
        let(:params) do
          {
            ensure: 'absent',
            version: '2013',
            edition: edition,
            sp: sp,
            license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
            products: ['Word'],
            deployment_root: '\\\\test-server\\packages'
          }
        end

        it do
          should contain_exec('uninstall-office').with(
            'command' => "& \"\\\\test-server\\packages\\OFFICE15\\#{edition}\\setup.exe\" /uninstall #{product} /config \"C:\\Windows\\Temp\\office_config.xml\"",
            'provider' => 'powershell',
            'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\15.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
        end
      end
    end
  end

  office_versions['2010']['editions'].keys.each do |edition|
    product = office_versions['2010']['editions'][edition]['office_product']

    describe "installing office 2010 #{edition}" do
      let :title do 'msoffice for 2010' end
      let(:params) do
        {
          version: '2010',
          edition: edition,
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\\\test-server\\packages'
        }
      end

      it do
        should contain_exec('install-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE14\\#{edition}\\x86\\setup.exe\" /config \"C:\\Windows\\Temp\\office_config.xml\"",
          'provider' => 'windows'
      )
      end

      it do
        should contain_exec('upgrade-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE14\\#{edition}\\x86\\setup.exe\" /modify #{product} /config \"C:\\Windows\\Temp\\office_config.xml\""
      )
      end
    end

    office_versions['2010']['service_packs'].keys.each do |sp|
      build = office_versions['2010']['service_packs'][sp]['build']
      describe "uninstalling office 2010 #{edition} SP#{sp}" do
        let :title do 'msoffice for 2010' end
        let(:params) do
          {
            ensure: 'absent',
            version: '2010',
            edition: edition,
            sp: sp,
            license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
            products: ['Word'],
            deployment_root: '\\\\test-server\\packages'
          }
        end

        it do
          should contain_exec('uninstall-office').with(
            'command' => "& \"\\\\test-server\\packages\\OFFICE14\\#{edition}\\x86\\setup.exe\" /uninstall #{product} /config \"C:\\Windows\\Temp\\office_config.xml\"",
            'provider' => 'powershell',
            'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\14.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
        end
      end
    end
  end

  office_versions['2007']['editions'].keys.each do |edition|
    product = office_versions['2007']['editions'][edition]['office_product']

    describe "installing office 2007 #{edition}" do
      let :title do 'msoffice for 2007' end
      let(:params) do
        {
          version: '2007',
          edition: edition,
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\\\test-server\\packages'
        }
      end

      it do
        should contain_exec('install-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE12\\#{edition}\\setup.exe\" /config \"C:\\Windows\\Temp\\office_config.xml\"",
          'provider' => 'windows'
      )
      end

      it do
        should contain_exec('upgrade-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE12\\#{edition}\\setup.exe\" /modify #{product} /config \"C:\\Windows\\Temp\\office_config.xml\""
      )
      end
    end

    office_versions['2007']['service_packs'].keys.each do |sp|
      build = office_versions['2007']['service_packs'][sp]['build']
      describe "uninstalling office 2007 #{edition}" do
        let :title do 'msoffice for 2007' end
        let(:params) do
          {
            ensure: 'absent',
            version: '2007',
            edition: edition,
            sp: sp,
            license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
            products: ['Word'],
            deployment_root: '\\\\test-server\\packages'
          }
        end

        it do
          should contain_exec('uninstall-office').with(
            'command' => "& \"\\\\test-server\\packages\\OFFICE12\\#{edition}\\setup.exe\" /uninstall #{product} /config \"C:\\Windows\\Temp\\office_config.xml\"",
            'provider' => 'powershell',
            'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\12.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
        end
      end
    end
  end

  office_versions['2003']['editions'].keys.each do |edition|
    product = office_versions['2003']['editions'][edition]['office_product']
    describe "installing office 2003 #{edition}" do
      let :title do 'msoffice for 2003' end
      let(:params) do
        {
          version: '2003',
          edition: edition,
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\\\test-server\\packages'
        }
      end

      it do
        should contain_exec('install-office').with(
          'command' => "\"\\\\test-server\\packages\\OFFICE11\\#{edition}\\SETUP.EXE\" /settings \"C:\\Windows\\Temp\\office_config.ini\"",
          'provider' => 'windows'
      )
      end
    end

    office_versions['2003']['service_packs'].keys.each do |sp|
      build = office_versions['2003']['service_packs'][sp]['build']
      describe "uninstalling office 2003 #{edition}" do
        let :title do 'msoffice for 2003' end
        let(:params) do
          {
            ensure: 'absent',
            version: '2003',
            sp: sp,
            edition: edition,
            license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
            products: ['Word'],
            deployment_root: '\\\\test-server\\packages'
          }
        end

        it do
          should contain_exec('uninstall-office').with(
            'command' => "& \"\\\\test-server\\packages\\OFFICE11\\#{edition}\\setup.exe\" /x #{product}.msi /qb",
            'provider' => 'powershell',
            'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\11.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )
        end
      end
    end
  end

  describe 'installing unknown version' do
    let :title do 'msoffice for unknown version' end
    let(:params) do
      {
        version: 'xxx',
        edition: 'Standard',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('install-office')
      end.to raise_error
    end
  end

  describe 'uninstalling unknown version' do
    let :title do 'msoffice for unknown version' end
    let(:params) do
      {
        ensure: 'absent',
        version: 'xxx',
        edition: 'Standard',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('uninstall-office')
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
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\\\test-server\\packages'
        }
      end

      it do
        expect do
          should contain_exec('install-office')
        end.to raise_error
      end
    end
  end

  describe 'incorrect license key' do
    let :title do 'msoffice with incorrect license key' end
    let(:params) do
      {
        version: '2010',
        edition: 'Standard',
        license_key: 'XXXXX-XXXXX-XXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('install-office')
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
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('install-office')
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
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('install-office')
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
          license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
          products: ['Word'],
          deployment_root: '\\\\test-server\\packages'
        }
      end

      it do
        expect do
          should contain_exec('install-office')
        end.not_to raise_error
      end
    end
  end

  describe 'invalid country' do
    let :title do 'msoffice with invalid countries' end
    let(:params) do
      {
        lang_code: 'fubar',
        version: '2010',
        edition: 'Standard',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('install-office')
      end.to raise_error
    end
  end

  describe 'incorrect sp' do
    let :title do 'msoffice with incorrect sp' end
    let(:params) do
      {
        sp: '5',
        version: '2010',
        edition: 'Standard',
        license_key: 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        products: ['Word'],
        deployment_root: '\\\\test-server\\packages'
      }
    end

    it do
      expect do
        should contain_exec('install-office')
      end.to raise_error
    end
  end
end
