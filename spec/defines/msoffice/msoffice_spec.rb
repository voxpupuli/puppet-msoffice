require 'spec_helper'

describe 'msoffice', :type => :define do

  $lcid_strings = {
      'af' => '1078', 'sq' => '1052', 'am' => '1118', 'ar-dz' => '5121', 'ar-bh' => '15361', 'ar-eg' => '3073', 'ar-iq' => '2049',
      'ar-jo' => '11265', 'ar-kw' => '13313', 'ar-lb' => '12289', 'ar-ly' => '4097', 'ar-ma' => '6145', 'ar-om' => '8193',
      'ar-qa' => '16385', 'ar-sa' => '1025', 'ar-sy' => '10241', 'ar-tn' => '7169', 'ar-ae' => '14337', 'ar-ye' => '9217',
      'hy' => '1067', 'as' => '1101', 'az-az' => '2092', 'eu' => '1069', 'be' => '1059', 'bn' => '2117', 'bs' => '5146',
      'bg' => '1026', 'my' => '1109', 'ca' => '1027', 'zh-cn' => '2052', 'zh-hk' => '3076', 'zh-mo' => '5124', 'zh-sg' => '4100',
      'zh-tw' => '1028', 'hr' => '1050', 'cs' => '1029', 'da' => '1030', 'dv' => '1125', 'nl-be' => '2067', 'nl-nl' => '1043',
      'en-au' => '3081', 'en-bz' => '10249', 'en-ca' => '4105', 'en-cb' => '9225', 'en-gb' => '2057', 'en-in' => '16393',
      'en-ie' => '6153', 'en-jm' => '8201', 'en-nz' => '5129', 'en-ph' => '13321', 'en-za' => '7177', 'en-tt' => '11273',
      'en-us' => '1033', 'et' => '1061', 'fo' => '1080', 'fa' => '1065', 'fi' => '1035', 'fr-be' => '2060', 'fr-ca' => '3084',
      'fr-fr' => '1036', 'fr-lu' => '5132', 'fr-ch' => '4108', 'mk' => '1071', 'gd-ie' => '2108', 'gd' => '1084', 'de-at' => '3079',
      'de-de' => '1031', 'de-li' => '5127', 'de-lu' => '4103', 'de-ch' => '2055', 'el' => '1032', 'gn' => '1140', 'gu' => '1095',
      'he' => '1037', 'hi' => '1081', 'hu' => '1038', 'is' => '1039', 'id' => '1057', 'it-it' => '1040', 'it-ch' => '2064',
      'ja' => '1041', 'kn' => '1099', 'ks' => '1120', 'kk' => '1087', 'km' => '1107', 'ko' => '1042', 'lo' => '1108', 'la' => '1142',
      'lv' => '1062', 'lt' => '1063', 'ms-bn' => '2110', 'ms-my' => '1086', 'ml' => '1100', 'mt' => '1082', 'mi' => '1153',
      'mr' => '1102', 'mn' => '2128', 'ne' => '1121', 'no-no' => '1044', 'or-or' => '1096', 'pl' => '1045', 'pt-br' => '1046',
      'pt-pt' => '2070', 'pa' => '1094', 'rm' => '1047', 'ro-mo' => '2072', 'ro' => '1048', 'ru' => '1049', 'ru-mo' => '2073',
      'sa' => '1103', 'sr-sp' => '3098', 'tn' => '1074', 'sd' => '1113', 'si' => '1115', 'sk' => '1051', 'sl' => '1060',
      'so' => '1143', 'sb' => '1070', 'es-ar' => '11274', 'es-bo' => '16394', 'es-cl' => '13322', 'es-co' => '9226', 'es-cr' => '5130',
      'es-do' => '7178', 'es-ec' => '12298', 'es-sv' => '17418', 'es-gt' => '4106', 'es-hn' => '18442', 'es-mx' => '2058',
      'es-ni' => '19466', 'es-pa' => '6154', 'es-py' => '15370', 'es-pe' => '10250', 'es-pr' => '20490', 'es-es' => '1034',
      'es-uy' => '14346', 'es-ve' => '8202', 'sw' => '1089', 'sv-fi' => '2077', 'sv-se' => '1053', 'tg' => '1064', 'ta' => '1097',
      'tt' => '1092', 'te' => '1098', 'th' => '1054', 'bo' => '1105', 'ts' => '1073', 'tr' => '1055', 'tk' => '1090', 'uk' => '1058',
      'ur' => '1056', 'uz-uz' => '2115', 'vi' => '1066', 'cy' => '1106', 'xh' => '1076', 'yi' => '1085', 'zu' => '1077'
  }

  describe "installing with unknown version" do
    let :title do "msoffice with unknown version" end
    let(:params) {{
      :version => 'xxx',
      :edition => 'Standard',
      :sp => '1',
      :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
      :deployment_root => '\\test-server\packages'
    }}

    it do
      expect {
        should contain_msoffice__package("microsoft office 2010")
      }.to raise_error
    end
  end

  ['2003','2007','2010','2013'].each do |version|
    describe "installing #{version} with wrong edition" do
      let :title do "msoffice for #{version}" end
      let(:params) {{
        :version => version,
        :edition => "fubar",
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it do
        expect {
          should contain_msoffice__package("microsoft office #{version}")
        }.to raise_error
      end
    end
  end

  describe "installing with unknown sp" do
    let :title do "msoffice unknown sp version" end
    let(:params) {{
      :version => '2010',
      :edition => 'Standard',
      :sp => '5',
      :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
      :deployment_root => '\\test-server\packages'
    }}

    it do
      expect {
        should contain_msoffice__package("microsoft office 2010")
      }.to raise_error
    end
  end

  describe "incorrect license key" do
    let :title do "msoffice with incorrect license key" end
    let(:params) {{
      :version => '2010',
      :edition => "Standard",
      :sp => '1',
      :license_key => 'XXXXX-XXXXX-XXXX-XXXXX-XXXXX',
      :products => ['Word'],
      :deployment_root => '\\test-server\packages'
    }}

    it do
      expect {
        should contain_msoffice__package("microsoft office 2010")
      }.to raise_error
    end
  end

  describe "incorrect arch" do
    let :title do "msoffice with incorrect arch" end
    let(:params) {{
      :arch => 'fubar',
      :version => '2010',
      :edition => "Standard",
      :sp => '1',
      :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
      :products => ['Word'],
      :deployment_root => '\\test-server\packages'
    }}

    it do
      expect {
        should contain_msoffice__package("microsoft office 2010")
      }.to raise_error
    end
  end

  describe "incorrect ensure" do
    let :title do "msoffice with incorrect ensure" end
    let(:params) {{
      :ensure => 'fubar',
      :version => '2010',
      :edition => "Standard",
      :sp => '1',
      :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
      :products => ['Word'],
      :deployment_root => '\\test-server\packages'
    }}

    it do
      expect {
        should contain_msoffice__package("microsoft office 2010")
      }.to raise_error
    end
  end

  $lcid_strings.keys.each do |lang_code|
    describe "valid country: #{lang_code}" do
      let :title do "msoffice with valid countries" end
      let(:params) {{
        :lang_code => lang_code,
        :version => '2010',
        :edition => "Standard",
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it do
        expect {
          should contain_msoffice__package("microsoft office 2010")
        }.to_not raise_error
      end
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "installs the core package for #{version}" do
      let :title do "office #{version}" end
      let(:params) {{
        :version => version,
        :edition => 'Standard',
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it { should contain_msoffice__package("microsoft office #{version}")}
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "installs the service pack when office #{version} present" do
      let :title do "office #{version}" end
      let(:params) {{
        :ensure => 'present',
        :version => version,
        :edition => 'Standard',
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it { should contain_msoffice__servicepack("microsoft office #{version} servicepack 1")}
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "does not install the service pack when office #{version} absent" do
      let :title do "office #{version}" end
      let(:params) {{
        :ensure => 'absent',
        :version => version,
        :edition => 'Standard',
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it { should_not contain_msoffice__servicepack("microsoft office #{version} servicepack 1")}
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "installs the lip when office #{version} present" do
      let :title do "office #{version}" end
      let(:params) {{
        :lang_code => 'de-de',
        :ensure => 'present',
        :version => version,
        :edition => 'Standard',
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it { should contain_msoffice__lip("microsoft lip de-de")}
    end
  end

  ['2003','2007','2010'].each do |version|
    describe "does not install the lip when office #{version} present" do
      let :title do "office #{version}" end
      let(:params) {{
        :lang_code => 'de-de',
        :ensure => 'absent',
        :version => version,
        :edition => 'Standard',
        :sp => '1',
        :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX',
        :products => ['Word'],
        :deployment_root => '\\test-server\packages'
      }}

      it { should_not contain_msoffice__lip("microsoft lip de-de")}
    end
  end

end
