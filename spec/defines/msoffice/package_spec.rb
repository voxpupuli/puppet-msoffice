require 'spec_helper'

describe 'msoffice::package', :type => :define do
  
  $temp_dir = 'C:\\Windows\\Temp'

  $office_versions = { "2003" => {
      "version" => "11",
      "editions" => {
          "Basic" => {
              "products" => ['Word','Excel','Outlook'],
              "office_product" => 'basic11',
          },
          "Student and Teacher" => {
              "products" => ['Word','Excel','Outlook','Powerpoint'],
              "office_product" => 'stdedu',
          },
          "Standard" => {
              "products" => ['Word','Excel','Outlook','Powerpoint'],
              "office_product" => 'std11',
          },
          "Small Business" => {
              "products" => ['Word','Excel','Outlook','Powerpoint','Publisher'],
              "office_product" => 'sbe11',
          },
          "Professional" => {
              "products" => ['Word','Excel','Outlook','Powerpoint','Publisher','InfoPath'],
              "office_product" => 'pro',
          }
      },
      "service_packs" => {
          "0" => {
              "build" => '11.0.5614.0'
          },
          "1" => {
              "setup" => 'Office2003SP1-kb842532-fullfile-enu',
              "build" => '11.0.6355.0'
          },
          "2" => {
              "setup" => 'Office2003SP2-KB887616-FullFile-ENU',
              "build" => '11.0.7969.0'
          },
          "3" => {
              "setup" => 'Office2003SP3-KB923618-FullFile-ENU',
              "build" => '11.0.8173.0'
          }
      }
  },
                       "2007" => {
                           "version" => "12",
                           "editions" => {
                               "Basic" => {
                                   "products" => ['Word','Excel','Outlook'],
                                   "office_product" => 'Basic',
                               },
                               "Home and Student" => {
                                   "products" => ['Word','Excel','Powerpoint'],
                                   "office_product" => 'Home and Student',
                               },
                               "Standard" => {
                                   "products" => ['Word','Excel','Powerpoint','Outlook'],
                                   "office_product" => 'Standard',
                               },
                               "Small Business" => {
                                   "products" => ['Word','Excel','Powerpoint','Outlook','Publisher'],
                                   "office_product" => 'Small Business',
                               },
                               "Professional" => {
                                   "products" => ['Word','Excel','Powerpoint','Outlook','Publisher','Access'],
                                   "office_product" => 'Pro',
                               },
                               "Professional Plus" => {
                                   "products" => ['Word','Excel','Powerpoint','Outlook','Publisher','Access','InfoPath','Communicator'],
                                   "office_product" => 'ProPlus',
                               },
                               "Ultimate" => {
                                   "products" => ['Word','Excel','Powerpoint','Outlook','Publisher','Access','InfoPath','Groove','OneNote'],
                                   "office_product" => 'Ultimater',
                               },
                               "Enterprise" => {
                                   "products" => ['Word','Excel','Powerpoint','Outlook','Publisher','Access','InfoPath','Communicator','Groove','OneNote'],
                                   "office_product" => 'Enterpise',
                               }
                           },
                           "service_packs" => {
                               "0" => {
                                   "build" => '12.0.4518.1014'
                               },
                               "1" => {
                                   "setup" => 'office2007sp1-kb936982-fullfile-en-us.exe',
                                   "build" => '12.0.6215.1000'
                               },
                               "2" => {
                                   "setup" => 'office2007sp2-kb953195-fullfile-en-us.exe',
                                   "build" => '12.0.6425.1000'
                               },
                               "3" => {
                                   "setup" => 'office2007sp3-kb2526086-fullfile-en-us.exe',
                                   "build" => '12.0.6612.1000'
                               }
                           }
                       },
                       "2010" => {
                           "version" => "14",
                           "editions" => {
                               "Starter" => {
                                   "products" => ['Word','Excel'],
                                   "office_product" => 'Starter',
                               },
                               "Personal" => {
                                   "products" => ['Word','Excel','Outlook'],
                                   "office_product" => 'Personal',
                               },
                               "Home and Student" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote'],
                                   "office_product" => 'Home and Student',
                               },
                               "Home and Business" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote','Outlook'],
                                   "office_product" => 'Home and Business',
                               },
                               "Standard" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote','Outlook','Publisher'],
                                   "office_product" => 'Standardr',
                               },
                               "Professional" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote','Outlook','Access'],
                                   "office_product" => 'Pro',
                               },
                               "Professional Plus" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote','Outlook','Access','InfoPath','Sharepoint Workspace'],
                                   "office_product" => 'ProPlus',
                               },
                           },
                           "service_packs" => {
                               "0" => {
                                   "build" => '14.0.4760.1000'
                               },
                               "1" => {
                                   "setup" => {
                                       "x86" => "officesuite2010sp1-kb2460049-x86-fullfile-en-us.exe",
                                       "x64" => "officesuite2010sp1-kb2460049-x64-fullfile-en-us.exe",
                                   },
                                   "build" => '14.0.6023.1000'
                               },
                               "2" => {
                                   "setup" => {
                                       "x86" => "officesp2010-kb2687455-fullfile-x86-en-us.exe",
                                       "x64" => "officesp2010-kb2687455-fullfile-x64-en-us.exe",
                                   },
                                   "build" => '14.0.7011.1000'
                               },
                           }
                       },
                       "2013" => {
                           "version" => "15",
                           "editions" => {
                               "Home and Student" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote'],
                                   "office_product" => 'Home and Student',
                               },
                               "Home and Business" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote', 'Outlook'],
                                   "office_product" => 'Home and Business',
                               },
                               "Standard" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote', 'Outlook','Publisher'],
                                   "office_product" => 'Standardr',
                               },
                               "Professional" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote', 'Outlook','Publisher','Access'],
                                   "office_product" => 'Pro',
                               },
                               "Professional Plus" => {
                                   "products" => ['Word','Excel','Powerpoint','OneNote', 'Outlook','Publisher','Access', 'InfoPath', 'Lync'],
                                   "office_product" => 'ProPlus',
                               }
                           },
                           "service_packs" => {
                               "0" => {
                                   "build" => '15.0.4420.1027'
                               },
                           }
                       }
  }

  $lcid_strings = ['af','sq','am','ar-dz','ar-bh','ar-eg','ar-iq','ar-jo','ar-kw','ar-lb','ar-ly','ar-ma','ar-om','ar-qa','ar-sa','ar-sy','ar-tn',
                   'ar-ae','ar-ye','hy','as','az-az','eu','be','bn','bn','bs','bg','my','ca','zh-cn','zh-hk','zh-mo','zh-sg','zh-tw','hr','cs','da',
                   'dv','nl-be','nl-nl','en-au','en-bz','en-ca','en-cb','en-gb','en-in','en-ie','en-jm','en-nz','en-ph','en-za','en-tt','en-us','et',
                   'fo','fa','fi','fr-be','fr-ca','fr-fr','fr-lu','fr-ch','mk','gd-ie','gd','de-at','de-de','de-li','de-lu','de-ch','el','gn','gu',
                   'he','hi','hu','is','id','it-it','it-ch','ja','kn','ks','kk','km','ko','lo','la','lv','lt','ms-bn','ms-my','ml','mt','mi','mr',
                   'mn','mn','ne','no-no','no-no','or','pl','pt-br','pt-pt','pa','rm','ro-mo','ro','ru','ru-mo','sa','sr-sp','sr-sp','tn','sd','si',
                   'sk','sl','so','sb','es-ar','es-bo','es-cl','es-co','es-cr','es-do','es-ec','es-sv','es-gt','es-hn','es-mx','es-ni','es-pa','es-py',
                   'es-pe','es-pr','es-es','es-uy','es-ve','sw','sv-fi','sv-se','tg','ta','tt','te','th','bo','ts','tr','tk','uk','ur','uz-uz','vi',
                   'cy','xh','yi','zu']

  
  let :hiera_data do
    { 
      :windows_deployment_root => '\\test-server\packages', 
      :company_name => 'Example Inc',
      :office_username => 'Joe'
    }
  end
  
  $office_versions['2013']['editions'].keys.each do |edition|
    product = $office_versions['2013']['editions'][edition]['office_product']
    
    describe "installing office 2013 #{edition}" do
      let :title do 'msoffice for 2013' end
      let :params do 
        { :version => '2013', :edition => edition, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should contain_exec('install-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE15\\#{edition}\\setup.exe\" /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
        'provider' => 'windows',
      ) }
      
      it { should contain_exec('upgrade-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE15\\#{edition}\\setup.exe\" /modify #{product} /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
      )}
    end

    $office_versions['2013']['service_packs'].keys.each do |sp|
      build = $office_versions['2013']['service_packs'][sp]['build']
      describe "uninstalling office 2013 #{edition} SP#{sp}" do
        let :title do 'msoffice for 2013' end
        let :params do
          { :ensure => 'absent', :version => '2013', :edition => edition, :sp => sp, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
        end

        it { should contain_exec('uninstall-office').with(
          'command' => "& \"\\test-server\\packages\\OFFICE15\\#{edition}\\setup.exe\" /uninstall #{product} /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\15.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        ) }
      end
    end
  end
  
  $office_versions['2010']['editions'].keys.each do |edition|
    product = $office_versions['2010']['editions'][edition]['office_product']
    
    describe "installing office 2010 #{edition}" do
      let :title do 'msoffice for 2010' end
      let :params do 
        { :version => '2010', :edition => edition, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should contain_exec('install-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE14\\#{edition}\\x86\\setup.exe\" /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
        'provider' => 'windows',
      ) }
      
      it { should contain_exec('upgrade-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE14\\#{edition}\\x86\\setup.exe\" /modify #{product} /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
      )}
    end

    $office_versions['2010']['service_packs'].keys.each do |sp|
      build = $office_versions['2010']['service_packs'][sp]['build']
      describe "uninstalling office 2010 #{edition} SP#{sp}" do
        let :title do 'msoffice for 2010' end
        let :params do
          { :ensure => 'absent', :version => '2010', :edition => edition, :sp => sp, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
        end

        it { should contain_exec('uninstall-office').with(
          'command' => "& \"\\test-server\\packages\\OFFICE14\\#{edition}\\x86\\setup.exe\" /uninstall #{product} /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\14.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )}
      end
    end
  end

  $office_versions['2007']['editions'].keys.each do |edition|
    product = $office_versions['2007']['editions'][edition]['office_product']
    
    describe "installing office 2007 #{edition}" do
      let :title do 'msoffice for 2007' end
      let :params do 
        { :version => '2007', :edition => edition, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should contain_exec('install-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE12\\#{edition}\\setup.exe\" /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
        'provider' => 'windows',
      ) }
      
      it { should contain_exec('upgrade-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE12\\#{edition}\\setup.exe\" /modify #{product} /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
      )}
    end

    $office_versions['2007']['service_packs'].keys.each do |sp|
      build = $office_versions['2007']['service_packs'][sp]['build']
      describe "uninstalling office 2007 #{edition}" do
        let :title do 'msoffice for 2007' end
        let :params do
          { :ensure => 'absent', :version => '2007', :edition => edition, :sp => sp, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
        end

        it { should contain_exec('uninstall-office').with(
          'command' => "& \"\\test-server\\packages\\OFFICE12\\#{edition}\\setup.exe\" /uninstall #{product} /config \"C:\\\\Windows\\\\Temp\\office_config.xml\"",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\12.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )}
      end
    end
  end
  
  $office_versions['2003']['editions'].keys.each do |edition|
    product = $office_versions['2003']['editions'][edition]['office_product']
    describe "installing office 2003 #{edition}" do
      let :title do 'msoffice for 2003' end
      let :params do 
        { :version => '2003', :edition => edition, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it { should contain_exec('install-office').with(
        'command' => "\"\\test-server\\packages\\OFFICE11\\#{edition}\\SETUP.EXE\" /settings \"C:\\\\Windows\\\\Temp\\office_config.ini\"",
        'provider' => 'windows',
      ) }
    end

    $office_versions['2003']['service_packs'].keys.each do |sp|
      build = $office_versions['2003']['service_packs'][sp]['build']
      describe "uninstalling office 2003 #{edition}" do
        let :title do 'msoffice for 2003' end
        let :params do
          { :ensure => 'absent', :version => '2003', :sp => sp, :edition => edition, :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
        end

        it { should contain_exec('uninstall-office').with(
          'command' => "& \"\\test-server\\packages\\OFFICE11\\#{edition}\\setup.exe\" /x #{product}.msi /qb",
          'provider' => 'powershell',
          'onlyif' => "if (Get-Item -LiteralPath \'\\HKLM:\\SOFTWARE\\Microsoft\\Office\\11.0\\Common\\ProductVersion\' -ErrorAction SilentlyContinue).GetValue(\'#{build}\')) { exit 1 }"
        )}
      end
    end
  end
  
  describe "installing unknown version" do
    let :title do "msoffice for unknown version" end
    let :params do 
      { :version => 'xxx', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('install-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The version agrument specified does not match a valid version of office' }
    end
  end
  
  describe "uninstalling unknown version" do
    let :title do "msoffice for unknown version" end
    let :params do 
      { :ensure => 'absent', :version => 'xxx', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('uninstall-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The version agrument specified does not match a valid version of office' }
    end
  end
  
  ['2003','2007','2010','2013'].each do |version|
    describe "installing #{version} with wrong edition" do
      let :title do "msoffice for #{version}" end
      let :params do 
        { :version => version, :edition => "fubar", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end
  
      it do
        expect {
          should contain_exec('install-office')
        }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The edition argument does not match a valid edition for the specified version of office' }
      end
    end
  end
  
  describe "incorrect license key" do
    let :title do "msoffice with incorrect license key" end
    let :params do 
      { :version => '2010', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('install-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The license_key argument speicifed is not correctly formatted' }
    end
  end
  
  describe "incorrect arch" do
    let :title do "msoffice with incorrect arch" end
    let :params do 
      { :arch => 'fubar', :version => '2010', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('install-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The arch argument specified does not match x86 or x64' }
    end
  end
  
  describe "incorrect ensure" do
    let :title do "msoffice with incorrect ensure" end
    let :params do 
      { :ensure => 'fubar', :version => '2010', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('install-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The ensure argument does not match present or absent' }
    end
  end

  $lcid_strings.each do |lang_code|
    describe "valid countries" do
      let :title do "msoffice with valid countries" end
      let :params do 
        { :lang_code => lang_code, :version => '2010', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
      end

      it do
        expect {
          should contain_exec('install-office')
        }.to_not raise_error(Puppet::Error)
      end
    end
  end
  
  describe "invalid country" do
    let :title do "msoffice with invalid countries" end
    let :params do 
      { :lang_code => 'fubar', :version => '2010', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('install-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The lang_code argument does not specifiy a valid language identifier' }
    end
  end

  describe "incorrect sp" do
    let :title do "msoffice with incorrect sp" end
    let :params do
      { :sp => '5', :version => '2010', :edition => "Standard", :license_key => 'XXXXX-XXXXX-XXXXX-XXXXX-XXXXX', :products => ['Word']}
    end

    it do
      expect {
        should contain_exec('install-office')
      }.to raise_error(Puppet::Error) {|e| expect(e.to_s).to match 'The service pack specified does not match 0-3' }
    end
  end
end