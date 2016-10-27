require 'spec_helper'

describe 'msoffice::params', type: :class do
  let :hiera_data do
    {
      windows_deployment_root: '\\test-server\packages',
      company_name: 'Example Inc',
      office_username: 'Joe'
    }
  end

  it { is_expected.to have_resource_count(0) }
end
