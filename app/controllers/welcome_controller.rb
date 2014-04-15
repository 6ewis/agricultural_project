class WelcomeController < ApplicationController

  def index
  	list
  end	

  def country
  end

  def valuechain
  end

  def gender_composition
  end

  def gender_leader
  end

  def list
  	@list_variables = Welcome.list_variables 
  end 
end
