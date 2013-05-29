class msoffice::params {
    $temp_dir = 'C:\\Windows\\Temp'
    $deployment_root = hiera('windows_deployment_root')
	$comp_name = hiera('company_name')
	$user_name = hiera('office_username')

    #Assumed RTM + 1033
	
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
  
}