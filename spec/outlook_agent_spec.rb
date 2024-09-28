require 'rails_helper'
require 'huginn_agent/spec_helper'

describe Agents::OutlookAgent do
  before(:each) do
    @valid_options = Agents::OutlookAgent.new.default_options
    @checker = Agents::OutlookAgent.new(:name => "OutlookAgent", :options => @valid_options)
    @checker.user = users(:bob)
    @checker.save!
  end

  pending "add specs here"
end
