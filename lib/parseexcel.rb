  file = File.dirname(__FILE__) + "/myspreadsheet.xlsx"
  s = Roo::Excelx.new(file)           # loads an Excel Spreadsheet

  class DataStructure

    def hash_by_country
       {
        :bangladesh => {},
        :ghana =>  {},
        :india =>  {},
        :malawi => {},
        :tanzania => {}
        }
    end

    def hash_by_value_chain
      {
      	:bangladesh_indigo => {},
      	:ghana_groundnuts => {},
      	:ghana_soy => {},
      	:india_maize => {},
      	:india_paddy => {},
      	:malawi_groundnut => {},
      	:malawi_soy => {},
      	:tanzania_cassava => {},
      	:tanzania_sesame => {}
      }
    end


    def hash_by 
     {
     #by_country
     :grade_by_country => hash_by_country, 
     :grade_in_percent_by_country => hash_by_country,
     :binary_by_country => hash_by_country,
     :binary_in_percent_by_country => hash_by_country,

     #by_value_chain
     :grade_by_value_chain => hash_by_value_chain,
     :grade_in_percent_by_value_chain => hash_by_value_chain,
     :binary_by_value_chain => hash_by_value_chain,
     :binary_in_percent_by_value_chain => hash_by_value_chain,
     }
    end

    def self.list_variables
        @list_variables 
    end

        @list_variables = [input_and_land_selection = {}, soil_and_water_management = {}, use_of_inputs = {} , planting = {},
        pest_and_disease_management = {} , spraying_and_pest_management = {}, harvest = {}, post_harvest_management = {},
        marketing = {}, record_keeping = {}, gender = {}]
  end

    

  DataStructure.list_variables.each_with_index do |variable, index|
    variable.merge!(DataStructure.new.hash_by)
    variable.merge!({:sheet => s.sheets[index]}) 	
  end


  #import excel data
  DataStructure.list_variables.each do |variable|
    #by_country
  	gc  = variable[:grade_by_country]
  	gpc = variable[:grade_in_percent_by_country]
  	bc  = variable[:binary_by_country]
  	bpc = variable[:binary_in_percent_by_country]

    #by_value_chain
    gvc  = variable[:grade_by_value_chain]
    gpvc = variable[:grade_in_percent_by_value_chain]
    bvc  = variable[:binary_by_value_chain]
    bpvc = variable[:binary_in_percent_by_value_chain]



  	s.default_sheet = variable[:sheet]
  	
    #by_country
    gc[:bangladesh][:A]     = s.cell('C', 3)
    gc[:bangladesh][:B]     = s.cell('C', 5)
    gc[:bangladesh][:C]     = s.cell('C', 7)
    gc[:bangladesh][:D]     = s.cell('C', 9)
    gc[:ghana][:A]          = s.cell('D', 3)
    gc[:ghana][:B]          = s.cell('D', 5)
    gc[:ghana][:C]          = s.cell('D', 7)
    gc[:ghana][:D]          = s.cell('D', 9)
    gc[:india][:A]          = s.cell('E', 3)
    gc[:india][:B]          = s.cell('E', 5)
    gc[:india][:C]          = s.cell('E', 7)
    gc[:india][:D]          = s.cell('E', 9)
    gc[:malawi][:A]         = s.cell('F', 3)
    gc[:malawi][:B]         = s.cell('F', 5)
    gc[:malawi][:C]         = s.cell('F', 7)
    gc[:malawi][:D]         = s.cell('F', 9)
    gc[:tanzania][:A]       = s.cell('G', 3)
    gc[:tanzania][:B]       = s.cell('G', 5)
    gc[:tanzania][:C]       = s.cell('G', 7)
    gc[:tanzania][:D]       = s.cell('G', 9)


    gpc[:bangladesh][:A]    = s.cell('C', 4)
    gpc[:bangladesh][:B]    = s.cell('C', 6)
    gpc[:bangladesh][:C]    = s.cell('C', 8)
    gpc[:bangladesh][:D]    = s.cell('C', 10)
    gpc[:ghana][:A]         = s.cell('D', 4)
    gpc[:ghana][:B]         = s.cell('D', 6)
    gpc[:ghana][:C]         = s.cell('D', 8)
    gpc[:ghana][:D]         = s.cell('D', 10)
    gpc[:india][:A]         = s.cell('E', 4)
    gpc[:india][:B]         = s.cell('E', 6)
    gpc[:india][:C]         = s.cell('E', 8)
    gpc[:india][:D]         = s.cell('E', 10)
    gpc[:malawi][:A]        = s.cell('F', 4)
    gpc[:malawi][:B]        = s.cell('F', 6)
    gpc[:malawi][:C]        = s.cell('F', 8)
    gpc[:malawi][:D]        = s.cell('F', 10)
    gpc[:tanzania][:A]      = s.cell('G', 4)
    gpc[:tanzania][:B]      = s.cell('G', 6)
    gpc[:tanzania][:C]      = s.cell('G', 8)
    gpc[:tanzania][:D]      = s.cell('G', 10)

    bc[:bangladesh][:No]    = s.cell('C', 15)
    bc[:bangladesh][:Yes]   = s.cell('C', 17)
    bc[:ghana][:No]         = s.cell('D', 15)
    bc[:ghana][:Yes]        = s.cell('D', 17)
    bc[:india][:No]         = s.cell('E', 15)
    bc[:india][:Yes]        = s.cell('E', 17)
    bc[:malawi][:No]        = s.cell('F', 15)
    bc[:malawi][:Yes]       = s.cell('F', 17)
    bc[:tanzania][:No]      = s.cell('G', 15)
    bc[:tanzania][:Yes]     = s.cell('G', 17)


    bpc[:bangladesh][:No]   = s.cell('C', 16)
    bpc[:bangladesh][:Yes]  = s.cell('C', 18)
    bpc[:ghana][:No]        = s.cell('D', 16)
    bpc[:ghana][:Yes]       = s.cell('D', 18)
    bpc[:india][:No]        = s.cell('E', 16)
    bpc[:india][:Yes]       = s.cell('E', 18)
    bpc[:malawi][:No]       = s.cell('F', 16)
    bpc[:malawi][:Yes]      = s.cell('F', 18)
    bpc[:tanzania][:No]     = s.cell('G', 16)
    bpc[:tanzania][:Yes]    = s.cell('G', 18)

    #by_value_chain


    gvc[:bangladesh_indigo][:A]      = s.cell('C', 26)
    gvc[:bangladesh_indigo][:B]      = s.cell('C', 28)
    gvc[:bangladesh_indigo][:C]      = s.cell('C', 30)
    gvc[:bangladesh_indigo][:D]      = s.cell('C', 32)
    gvc[:ghana_groundnuts][:A]       = s.cell('D', 26)
    gvc[:ghana_groundnuts][:B]       = s.cell('D', 28)
    gvc[:ghana_groundnuts][:C]       = s.cell('D', 30)
    gvc[:ghana_groundnuts][:D]       = s.cell('D', 32)
    gvc[:ghana_soy][:A]              = s.cell('E', 26)
    gvc[:ghana_soy][:B]              = s.cell('E', 28)
    gvc[:ghana_soy][:C]              = s.cell('E', 30)
    gvc[:ghana_soy][:D]              = s.cell('E', 32)
    gvc[:india_maize][:A]            = s.cell('F', 26)
    gvc[:india_maize][:B]            = s.cell('F', 28)
    gvc[:india_maize][:C]            = s.cell('F', 30)
    gvc[:india_maize][:D]            = s.cell('F', 32)
    gvc[:india_paddy][:A]            = s.cell('G', 26)
    gvc[:india_paddy][:B]            = s.cell('G', 28)
    gvc[:india_paddy][:C]            = s.cell('G', 30)
    gvc[:india_paddy][:D]            = s.cell('G', 32)
    gvc[:malawi_groundnut][:A]       = s.cell('H', 26)
    gvc[:malawi_groundnut][:B]       = s.cell('H', 28)
    gvc[:malawi_groundnut][:C]       = s.cell('H', 30)
    gvc[:malawi_groundnut][:D]       = s.cell('H', 32)
    gvc[:malawi_soy][:A]             = s.cell('I', 26)
    gvc[:malawi_soy][:B]             = s.cell('I', 28)
    gvc[:malawi_soy][:C]             = s.cell('I', 30)
    gvc[:malawi_soy][:D]             = s.cell('I', 32)
    gvc[:tanzania_cassava][:A]       = s.cell('J', 26)
    gvc[:tanzania_cassava][:B]       = s.cell('J', 28)
    gvc[:tanzania_cassava][:C]       = s.cell('J', 30)
    gvc[:tanzania_cassava][:D]       = s.cell('J', 32)
    gvc[:tanzania_sesame][:A]        = s.cell('K', 26)
    gvc[:tanzania_sesame][:B]        = s.cell('K', 28)
    gvc[:tanzania_sesame][:C]        = s.cell('K', 30)
    gvc[:tanzania_sesame][:D]        = s.cell('K', 32)

    gpvc[:bangladesh_indigo][:A]     = s.cell('C', 27)
    gpvc[:bangladesh_indigo][:B]     = s.cell('C', 29)
    gpvc[:bangladesh_indigo][:C]     = s.cell('C', 31)
    gpvc[:bangladesh_indigo][:D]     = s.cell('C', 33)
    gpvc[:ghana_groundnuts][:A]      = s.cell('D', 27)
    gpvc[:ghana_groundnuts][:B]      = s.cell('D', 29)
    gpvc[:ghana_groundnuts][:C]      = s.cell('D', 31)
    gpvc[:ghana_groundnuts][:D]      = s.cell('D', 33)
    gpvc[:ghana_soy][:A]             = s.cell('E', 27)
    gpvc[:ghana_soy][:B]             = s.cell('E', 29)
    gpvc[:ghana_soy][:C]             = s.cell('E', 31)
    gpvc[:ghana_soy][:D]             = s.cell('E', 33)
    gpvc[:india_maize][:A]           = s.cell('F', 27)
    gpvc[:india_maize][:B]           = s.cell('F', 29)
    gpvc[:india_maize][:C]           = s.cell('F', 31)
    gpvc[:india_maize][:D]           = s.cell('F', 33)
    gpvc[:india_paddy][:A]           = s.cell('G', 27)
    gpvc[:india_paddy][:B]           = s.cell('G', 29)
    gpvc[:india_paddy][:C]           = s.cell('G', 31)
    gpvc[:india_paddy][:D]           = s.cell('G', 33)
    gpvc[:malawi_groundnut][:A]      = s.cell('H', 27)
    gpvc[:malawi_groundnut][:B]      = s.cell('H', 29)
    gpvc[:malawi_groundnut][:C]      = s.cell('H', 31)
    gpvc[:malawi_groundnut][:D]      = s.cell('H', 33)
    gpvc[:malawi_soy][:A]            = s.cell('I', 27)
    gpvc[:malawi_soy][:B]            = s.cell('I', 29)
    gpvc[:malawi_soy][:C]            = s.cell('I', 31)
    gpvc[:malawi_soy][:D]            = s.cell('I', 33)
    gpvc[:tanzania_cassava][:A]      = s.cell('J', 27)
    gpvc[:tanzania_cassava][:B]      = s.cell('J', 29)
    gpvc[:tanzania_cassava][:C]      = s.cell('J', 31)
    gpvc[:tanzania_cassava][:D]      = s.cell('J', 33)
    gpvc[:tanzania_sesame][:A]       = s.cell('K', 27)
    gpvc[:tanzania_sesame][:B]       = s.cell('K', 29)
    gpvc[:tanzania_sesame][:C]       = s.cell('K', 31)
    gpvc[:tanzania_sesame][:D]       = s.cell('K', 33)


    bvc[:bangladesh_indigo][:No]     = s.cell('C', 38)
    bvc[:bangladesh_indigo][:Yes]    = s.cell('C', 40)
    bvc[:ghana_groundnuts][:No]      = s.cell('D', 38)
    bvc[:ghana_groundnuts][:Yes]     = s.cell('D', 40)
    bvc[:ghana_soy][:No]             = s.cell('E', 38)
    bvc[:ghana_soy][:Yes]            = s.cell('E', 40)
    bvc[:india_maize][:No]           = s.cell('F', 38)
    bvc[:india_maize][:Yes]          = s.cell('F', 40)
    bvc[:india_paddy][:No]           = s.cell('G', 38)
    bvc[:india_paddy][:Yes]          = s.cell('G', 40)
    bvc[:malawi_groundnut][:No]      = s.cell('H', 38)
    bvc[:malawi_groundnut][:Yes]     = s.cell('H', 40)
    bvc[:malawi_soy][:No]            = s.cell('I', 38)
    bvc[:malawi_soy][:Yes]           = s.cell('I', 40)
    bvc[:tanzania_cassava][:No]      = s.cell('J', 38)
    bvc[:tanzania_cassava][:Yes]     = s.cell('J', 40)
    bvc[:tanzania_sesame][:No]       = s.cell('K', 38)
    bvc[:tanzania_sesame][:Yes]      = s.cell('K', 40)

    bpvc[:bangladesh_indigo][:No]    = s.cell('C', 39)
    bpvc[:bangladesh_indigo][:Yes]   = s.cell('C', 41)
    bpvc[:ghana_groundnuts][:No]     = s.cell('D', 39)
    bpvc[:ghana_groundnuts][:Yes]    = s.cell('D', 41)
    bpvc[:ghana_soy][:No]            = s.cell('E', 39)
    bpvc[:ghana_soy][:Yes]           = s.cell('E', 41)
    bpvc[:india_maize][:No]          = s.cell('F', 39)
    bpvc[:india_maize][:Yes]         = s.cell('F', 41)
    bpvc[:india_paddy][:No]          = s.cell('G', 39)
    bpvc[:india_paddy][:Yes]         = s.cell('G', 41)
    bpvc[:malawi_groundnut][:No]     = s.cell('H', 39)
    bpvc[:malawi_groundnut][:Yes]    = s.cell('H', 41)
    bpvc[:malawi_soy][:No]           = s.cell('I', 39)
    bpvc[:malawi_soy][:Yes]          = s.cell('I', 41)
    bpvc[:tanzania_cassava][:No]     = s.cell('J', 39)
    bpvc[:tanzania_cassava][:Yes]    = s.cell('J', 41)
    bpvc[:tanzania_sesame][:No]      = s.cell('K', 39)
    bpvc[:tanzania_sesame][:Yes]     = s.cell('K', 41)
   end

