require 'spec_helper'

describe 'msoffice::lip', type: :define do
  describe 'installing with unknown version' do
    let :title do
      'lip with unknown version'
    end
    let(:params) do
      {
        version: 'xxx',
        lang_code: 'en-us',
        arch: 'x86',
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect { is_expected.to contain_exec('install-lip') }.
        to raise_error(Puppet::Error, %r{The version argument specified does not match a valid version of office})
    end
  end

  describe 'incorrect arch' do
    let :title do
      'lip with incorrect arch'
    end
    let(:params) do
      {
        version: '2010',
        lang_code: 'en-us',
        arch: 'fubar',
        deployment_root: '\\test-server\packages'
      }
    end

    it do
      expect { is_expected.to contain_exec('install-lip') }.
        to raise_error(Puppet::Error, %r{The arch argument specified does not match x86 or x64})
    end
  end

  ['de-de'].each do |lang_code|
    lip_root = '\\test-server\\packages\\OFFICE14\\LIPs'
    lip_reg_key = 'HKLM:\\SOFTWARE\\Microsoft\\Office\\14.0\\Common\\LanguageResources\\InstalledUIs'
    setup = "languageinterfacepack-x86-#{lang_code}.exe"
    describe "valid country: #{lang_code}" do
      let :title do
        'lip with valid countries'
      end
      let(:params) do
        {
          version: '2010',
          lang_code: lang_code,
          arch: 'x86',
          deployment_root: '\\test-server\packages'
        }
      end

      it do
        is_expected.to contain_exec('install-lip').with(
          'command' => "& \"#{lip_root}\\#{setup}\" /q /norestart",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\#{lip_reg_key}\' -ErrorAction SilentlyContinue).GetValue(\'1031\')) { exit 1 }"
        )
      end
    end
  end
end
