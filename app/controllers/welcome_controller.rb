class WelcomeController < ApplicationController

  def index
  	list
  end	

  def country
  end

  def valuechain
  end

  def group_composition
  end

  def group_leader
  end

  def list
  	@list_variables = Welcome.list_variables 
  end 
end
