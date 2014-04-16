  file = File.dirname(__FILE__) + "/myspreadsheet2.xlsx"
  s = Roo::Excelx.new(file)           # loads an Excel Spreadsheet

  class DataStructure

    def hash_by_country
       {
        :bangladesh => {},
        :ghana =>  {},
        :india =>  {},
        :malawi => {},
        :tanzania => {},
        :overall => {}
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
        :tanzania_sesame => {},
        :overall => {}
      }
    end

    def hash_by_group_composition
        {
            :bangladesh_female_groups => {},
            :bangladesh_mixed_groups => {},
            :ghana_female_groups => {},
            :ghana_mixed_groups => {},
            :india_female_groups => {},
            :india_mixed_groups => {},
            :malawi_female_groups => {},
            :malawi_mixed_groups => {},
            :tanzania_female_groups => {},
            :tanzania_mixed_groups => {},
            :total_female_groups => {},
            :total_mixed_groups => {}
        }
    end

    def hash_by_group_leader
        {
            :bangladesh_female_group_leader => {},
            :bangladesh_male_group_leader => {},
            :bangladesh_mixed_group_leader => {},
            :ghana_female_group_leader => {},
            :ghana_male_group_leader => {},
            :ghana_mixed_group_leader => {},
            :india_female_group_leader => {},
            :india_male_group_leader => {},
            :india_mixed_group_leader => {},
            :malawi_female_group_leader => {},
            :malawi_male_group_leader => {},
            :malawi_mixed_group_leader => {},
            :tanzania_female_group_leader => {},
            :tanzania_male_group_leader => {},
            :tanzania_mixed_group_leader => {},
            :total_female_group_leader => {},
            :total_male_group_leader => {},
            :total_mixed_group_leader => {}
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

     #by_group_composition
     :grade_by_group_composition => hash_by_group_composition,
     :grade_in_percent_by_group_composition => hash_by_group_composition,
     :binary_by_group_composition => hash_by_group_composition,
     :binary_in_percent_by_group_composition => hash_by_group_composition,

     #by_group_leader
     :grade_by_group_leader => hash_by_group_leader,
     :grade_in_percent_by_group_leader => hash_by_group_leader,
     :binary_by_group_leader => hash_by_group_leader,
     :binary_in_percent_by_group_leader => hash_by_group_leader,
     
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

    #by_group_composition
    ggc  = variable[:grade_by_group_composition]
    gpgc = variable[:grade_in_percent_by_group_composition]
    bgc  = variable[:binary_by_group_composition]
    bpgc = variable[:binary_in_percent_by_group_composition]

    #by_group_leader
    ggl  = variable[:grade_by_group_leader]
    gpgl = variable[:grade_in_percent_by_group_leader]
    bgl  = variable[:binary_by_group_leader]
    bpgl = variable[:binary_in_percent_by_group_leader]

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
    gc[:overall][:A]        = s.cell('H', 3)
    gc[:overall][:B]        = s.cell('H', 5)
    gc[:overall][:C]        = s.cell('H', 7)
    gc[:overall][:D]        = s.cell('H', 9)


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
    gpc[:overall][:A]       = s.cell('H', 4)
    gpc[:overall][:B]       = s.cell('H', 6)
    gpc[:overall][:C]       = s.cell('H', 8)
    gpc[:overall][:D]       = s.cell('H', 10)

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
    bc[:overall][:No]       = s.cell('H', 15)
    bc[:overall][:Yes]      = s.cell('H', 17)


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
    bpc[:overall][:No]      = s.cell('H', 16)
    bpc[:overall][:Yes]     = s.cell('H', 18)

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
    gvc[:overall][:A]                = s.cell('L', 26)
    gvc[:overall][:B]                = s.cell('L', 28)
    gvc[:overall][:C]                = s.cell('L', 30)
    gvc[:overall][:D]                = s.cell('L', 32)

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
    gpvc[:overall][:A]               = s.cell('L', 27)
    gpvc[:overall][:B]               = s.cell('L', 29)
    gpvc[:overall][:C]               = s.cell('L', 31)
    gpvc[:overall][:D]               = s.cell('L', 33)


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
    bvc[:overall][:No]               = s.cell('L', 38)
    bvc[:overall][:Yes]              = s.cell('L', 40)

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
    bpvc[:overall][:No]              = s.cell('L', 39)
    bpvc[:overall][:Yes]             = s.cell('L', 41)

    #by_group_composition

    ggc[:bangladesh_female_groups][:A] = s.cell('C', 48)
    ggc[:bangladesh_female_groups][:B] = s.cell('C', 50)
    ggc[:bangladesh_female_groups][:C] = s.cell('C', 52)
    ggc[:bangladesh_female_groups][:D] = s.cell('C', 54)
    ggc[:bangladesh_mixed_groups][:A]  = s.cell('D', 48)
    ggc[:bangladesh_mixed_groups][:B]  = s.cell('D', 50)
    ggc[:bangladesh_mixed_groups][:C]  = s.cell('D', 52)
    ggc[:bangladesh_mixed_groups][:D]  = s.cell('D', 54)
    ggc[:ghana_female_groups][:A]      = s.cell('E', 48) 
    ggc[:ghana_female_groups][:B]      = s.cell('E', 50)
    ggc[:ghana_female_groups][:C]      = s.cell('E', 52)
    ggc[:ghana_female_groups][:D]      = s.cell('E', 54)
    ggc[:ghana_mixed_groups][:A]       = s.cell('F', 48)
    ggc[:ghana_mixed_groups][:B]       = s.cell('F', 50)
    ggc[:ghana_mixed_groups][:C]       = s.cell('F', 52)
    ggc[:ghana_mixed_groups][:D]       = s.cell('F', 54)
    ggc[:india_female_groups][:A]      = s.cell('G', 48)
    ggc[:india_female_groups][:B]      = s.cell('G', 50)
    ggc[:india_female_groups][:C]      = s.cell('G', 52)
    ggc[:india_female_groups][:D]      = s.cell('G', 54)
    ggc[:india_mixed_groups][:A]       = s.cell('H', 48)
    ggc[:india_mixed_groups][:B]       = s.cell('H', 50)
    ggc[:india_mixed_groups][:C]       = s.cell('H', 52)
    ggc[:india_mixed_groups][:D]       = s.cell('H', 54)
    ggc[:malawi_female_groups][:A]     = s.cell('I', 48)
    ggc[:malawi_female_groups][:B]     = s.cell('I', 50)
    ggc[:malawi_female_groups][:C]     = s.cell('I', 52)
    ggc[:malawi_female_groups][:D]     = s.cell('I', 54)
    ggc[:malawi_mixed_groups][:A]      = s.cell('J', 48)
    ggc[:malawi_mixed_groups][:B]      = s.cell('J', 50)
    ggc[:malawi_mixed_groups][:C]      = s.cell('J', 52)
    ggc[:malawi_mixed_groups][:D]      = s.cell('J', 54)
    ggc[:tanzania_female_groups][:A]   = s.cell('K', 48)
    ggc[:tanzania_female_groups][:B]   = s.cell('K', 50)
    ggc[:tanzania_female_groups][:C]   = s.cell('K', 52)
    ggc[:tanzania_female_groups][:D]   = s.cell('K', 54)
    ggc[:tanzania_mixed_groups][:A]    = s.cell('L', 48)
    ggc[:tanzania_mixed_groups][:B]    = s.cell('L', 50)
    ggc[:tanzania_mixed_groups][:C]    = s.cell('L', 52)  
    ggc[:tanzania_mixed_groups][:D]    = s.cell('L', 54)
    ggc[:total_female_groups][:A]      = s.cell('M', 48)
    ggc[:total_female_groups][:B]      = s.cell('M', 50)
    ggc[:total_female_groups][:C]      = s.cell('M', 52)
    ggc[:total_female_groups][:D]      = s.cell('M', 54)
    ggc[:total_mixed_groups][:A]       = s.cell('N', 48)
    ggc[:total_mixed_groups][:B]       = s.cell('N', 50)
    ggc[:total_mixed_groups][:C]       = s.cell('N', 52)
    ggc[:total_mixed_groups][:D]       = s.cell('N', 54)

    gpgc[:bangladesh_female_groups][:A] = s.cell('C', 49)
    gpgc[:bangladesh_female_groups][:B] = s.cell('C', 51)
    gpgc[:bangladesh_female_groups][:C] = s.cell('C', 53)
    gpgc[:bangladesh_female_groups][:D] = s.cell('C', 55)
    gpgc[:bangladesh_mixed_groups][:A]  = s.cell('D', 49)
    gpgc[:bangladesh_mixed_groups][:B]  = s.cell('D', 51)
    gpgc[:bangladesh_mixed_groups][:C]  = s.cell('D', 53)
    gpgc[:bangladesh_mixed_groups][:D]  = s.cell('D', 55)
    gpgc[:ghana_female_groups][:A]      = s.cell('E', 49) 
    gpgc[:ghana_female_groups][:B]      = s.cell('E', 51)
    gpgc[:ghana_female_groups][:C]      = s.cell('E', 53)
    gpgc[:ghana_female_groups][:D]      = s.cell('E', 55)
    gpgc[:ghana_mixed_groups][:A]       = s.cell('F', 49)
    gpgc[:ghana_mixed_groups][:B]       = s.cell('F', 51)
    gpgc[:ghana_mixed_groups][:C]       = s.cell('F', 53)
    gpgc[:ghana_mixed_groups][:D]       = s.cell('F', 55)
    gpgc[:india_female_groups][:A]      = s.cell('G', 49)
    gpgc[:india_female_groups][:B]      = s.cell('G', 51)
    gpgc[:india_female_groups][:C]      = s.cell('G', 53)
    gpgc[:india_female_groups][:D]      = s.cell('G', 55)
    gpgc[:india_mixed_groups][:A]       = s.cell('H', 49)
    gpgc[:india_mixed_groups][:B]       = s.cell('H', 51)
    gpgc[:india_mixed_groups][:C]       = s.cell('H', 53)
    gpgc[:india_mixed_groups][:D]       = s.cell('H', 55)
    gpgc[:malawi_female_groups][:A]     = s.cell('I', 49)
    gpgc[:malawi_female_groups][:B]     = s.cell('I', 51)
    gpgc[:malawi_female_groups][:C]     = s.cell('I', 53)
    gpgc[:malawi_female_groups][:D]     = s.cell('I', 55)
    gpgc[:malawi_mixed_groups][:A]      = s.cell('J', 49)
    gpgc[:malawi_mixed_groups][:B]      = s.cell('J', 51)
    gpgc[:malawi_mixed_groups][:C]      = s.cell('J', 53)
    gpgc[:malawi_mixed_groups][:D]      = s.cell('J', 55)
    gpgc[:tanzania_female_groups][:A]   = s.cell('K', 49)
    gpgc[:tanzania_female_groups][:B]   = s.cell('K', 51)
    gpgc[:tanzania_female_groups][:C]   = s.cell('K', 53)
    gpgc[:tanzania_female_groups][:D]   = s.cell('K', 55)
    gpgc[:tanzania_mixed_groups][:A]    = s.cell('L', 49)
    gpgc[:tanzania_mixed_groups][:B]    = s.cell('L', 51)
    gpgc[:tanzania_mixed_groups][:C]    = s.cell('L', 53)  
    gpgc[:tanzania_mixed_groups][:D]    = s.cell('L', 55)
    gpgc[:total_female_groups][:A]      = s.cell('M', 49)
    gpgc[:total_female_groups][:B]      = s.cell('M', 51)
    gpgc[:total_female_groups][:C]      = s.cell('M', 53)
    gpgc[:total_female_groups][:D]      = s.cell('M', 55)
    gpgc[:total_mixed_groups][:A]       = s.cell('N', 49)
    gpgc[:total_mixed_groups][:B]       = s.cell('N', 51)
    gpgc[:total_mixed_groups][:C]       = s.cell('N', 53)
    gpgc[:total_mixed_groups][:D]       = s.cell('N', 55)

    bgc[:bangladesh_female_groups][:No] = s.cell('B', 62)
    bgc[:bangladesh_female_groups][:Yes]= s.cell('B', 64)
    bgc[:bangladesh_mixed_groups][:No]  = s.cell('C', 62)
    bgc[:bangladesh_mixed_groups][:Yes] = s.cell('C', 64)
    bgc[:ghana_female_groups][:No]      = s.cell('D', 62)
    bgc[:ghana_female_groups][:Yes]     = s.cell('D', 64)
    bgc[:ghana_mixed_groups][:No]       = s.cell('E', 62)
    bgc[:ghana_mixed_groups][:Yes]      = s.cell('E', 64)
    bgc[:india_female_groups][:No]      = s.cell('F', 62) 
    bgc[:india_female_groups][:Yes]     = s.cell('F', 64)
    bgc[:india_mixed_groups][:No]       = s.cell('G', 62) 
    bgc[:india_mixed_groups][:Yes]      = s.cell('G', 64)
    bgc[:malawi_female_groups][:No]     = s.cell('H', 62)
    bgc[:malawi_female_groups][:Yes]    = s.cell('H', 64)
    bgc[:malawi_mixed_groups][:No]      = s.cell('I', 62)
    bgc[:malawi_mixed_groups][:Yes]     = s.cell('I', 64)
    bgc[:tanzania_female_groups][:No]   = s.cell('J', 62)
    bgc[:tanzania_female_groups][:Yes]  = s.cell('J', 64)
    bgc[:tanzania_mixed_groups][:No]    = s.cell('K', 62)
    bgc[:tanzania_mixed_groups][:Yes]   = s.cell('K', 64)
    bgc[:total_female_groups][:No]      = s.cell('L', 62)
    bgc[:total_female_groups][:Yes]     = s.cell('L', 64)
    bgc[:total_mixed_groups][:No]       = s.cell('M', 62)
    bgc[:total_mixed_groups][:Yes]      = s.cell('M', 64) 

    bpgc[:bangladesh_female_groups][:No] = s.cell('B', 63)
    bpgc[:bangladesh_female_groups][:Yes]= s.cell('B', 65)
    bpgc[:bangladesh_mixed_groups][:No]  = s.cell('C', 63)
    bpgc[:bangladesh_mixed_groups][:Yes] = s.cell('C', 65)
    bpgc[:ghana_female_groups][:No]      = s.cell('D', 63)
    bpgc[:ghana_female_groups][:Yes]     = s.cell('D', 65)
    bpgc[:ghana_mixed_groups][:No]       = s.cell('E', 63)
    bpgc[:ghana_mixed_groups][:Yes]      = s.cell('E', 65)
    bpgc[:india_female_groups][:No]      = s.cell('F', 63) 
    bpgc[:india_female_groups][:Yes]     = s.cell('F', 65)
    bpgc[:india_mixed_groups][:No]       = s.cell('G', 63) 
    bpgc[:india_mixed_groups][:Yes]      = s.cell('G', 65)
    bpgc[:malawi_female_groups][:No]     = s.cell('H', 63)
    bpgc[:malawi_female_groups][:Yes]    = s.cell('H', 65)
    bpgc[:malawi_mixed_groups][:No]      = s.cell('I', 63)
    bpgc[:malawi_mixed_groups][:Yes]     = s.cell('I', 65)
    bpgc[:tanzania_female_groups][:No]   = s.cell('J', 63)
    bpgc[:tanzania_female_groups][:Yes]  = s.cell('J', 65)
    bpgc[:tanzania_mixed_groups][:No]    = s.cell('K', 63)
    bpgc[:tanzania_mixed_groups][:Yes]   = s.cell('K', 65)
    bpgc[:total_female_groups][:No]      = s.cell('L', 63)
    bpgc[:total_female_groups][:Yes]     = s.cell('L', 65)
    bpgc[:total_mixed_groups][:No]       = s.cell('M', 63)
    bpgc[:total_mixed_groups][:Yes]      = s.cell('M', 65) 

    #by_group_leader

    ggl[:bangladesh_female_group_leader][:A] = s.cell('C', 73)
    ggl[:bangladesh_female_group_leader][:B] = s.cell('C', 75)
    ggl[:bangladesh_female_group_leader][:C] = s.cell('C', 77)
    ggl[:bangladesh_female_group_leader][:D] = s.cell('C', 79)
    ggl[:bangladesh_male_group_leader][:A]   = s.cell('D', 73)
    ggl[:bangladesh_male_group_leader][:B]   = s.cell('D', 75)
    ggl[:bangladesh_male_group_leader][:C]   = s.cell('D', 77)
    ggl[:bangladesh_male_group_leader][:D]   = s.cell('D', 79)
    ggl[:bangladesh_mixed_group_leader][:A]  = s.cell('E', 73)
    ggl[:bangladesh_mixed_group_leader][:B]  = s.cell('E', 75)
    ggl[:bangladesh_mixed_group_leader][:C]  = s.cell('E', 77)
    ggl[:bangladesh_mixed_group_leader][:D]  = s.cell('E', 79)
    ggl[:ghana_female_group_leader][:A]      = s.cell('F', 73)
    ggl[:ghana_female_group_leader][:B]      = s.cell('F', 75)
    ggl[:ghana_female_group_leader][:C]      = s.cell('F', 77)
    ggl[:ghana_female_group_leader][:D]      = s.cell('F', 79)
    ggl[:ghana_male_group_leader][:A]        = s.cell('G', 73)
    ggl[:ghana_male_group_leader][:B]        = s.cell('G', 75)
    ggl[:ghana_male_group_leader][:C]        = s.cell('G', 77)
    ggl[:ghana_male_group_leader][:D]        = s.cell('G', 79)
    ggl[:ghana_mixed_group_leader][:A]       = s.cell('H', 73)
    ggl[:ghana_mixed_group_leader][:B]       = s.cell('H', 75)
    ggl[:ghana_mixed_group_leader][:C]       = s.cell('H', 77)
    ggl[:ghana_mixed_group_leader][:D]       = s.cell('H', 79)
    ggl[:india_female_group_leader][:A]      = s.cell('I', 73)
    ggl[:india_female_group_leader][:B]      = s.cell('I', 75)
    ggl[:india_female_group_leader][:C]      = s.cell('I', 77)
    ggl[:india_female_group_leader][:D]      = s.cell('I', 79)
    ggl[:india_male_group_leader][:A]        = s.cell('J', 73)
    ggl[:india_male_group_leader][:B]        = s.cell('J', 75)
    ggl[:india_male_group_leader][:C]        = s.cell('J', 77)
    ggl[:india_male_group_leader][:D]        = s.cell('J', 79)
    ggl[:india_mixed_group_leader][:A]       = s.cell('K', 73)
    ggl[:india_mixed_group_leader][:B]       = s.cell('K', 75)
    ggl[:india_mixed_group_leader][:C]       = s.cell('K', 77)
    ggl[:india_mixed_group_leader][:D]       = s.cell('K', 79)
    ggl[:malawi_female_group_leader][:A]     = s.cell('L', 73)
    ggl[:malawi_female_group_leader][:B]     = s.cell('L', 75)
    ggl[:malawi_female_group_leader][:C]     = s.cell('L', 77)
    ggl[:malawi_female_group_leader][:D]     = s.cell('L', 79)
    ggl[:malawi_male_group_leader][:A]       = s.cell('M', 73)
    ggl[:malawi_male_group_leader][:B]       = s.cell('M', 75)
    ggl[:malawi_male_group_leader][:C]       = s.cell('M', 77)
    ggl[:malawi_male_group_leader][:D]       = s.cell('M', 79)
    ggl[:malawi_mixed_group_leader][:A]      = s.cell('N', 73)
    ggl[:malawi_mixed_group_leader][:B]      = s.cell('N', 75)
    ggl[:malawi_mixed_group_leader][:C]      = s.cell('N', 77)
    ggl[:malawi_mixed_group_leader][:D]      = s.cell('N', 79)
    ggl[:tanzania_female_group_leader][:A]   = s.cell('O', 73)
    ggl[:tanzania_female_group_leader][:B]   = s.cell('O', 75)
    ggl[:tanzania_female_group_leader][:C]   = s.cell('O', 77)
    ggl[:tanzania_female_group_leader][:D]   = s.cell('O', 79)
    ggl[:tanzania_male_group_leader][:A]     = s.cell('P', 73)
    ggl[:tanzania_male_group_leader][:B]     = s.cell('P', 75)
    ggl[:tanzania_male_group_leader][:C]     = s.cell('P', 77)
    ggl[:tanzania_male_group_leader][:D]     = s.cell('P', 79)
    ggl[:tanzania_mixed_group_leader][:A]    = s.cell('Q', 73)
    ggl[:tanzania_mixed_group_leader][:B]    = s.cell('Q', 75)
    ggl[:tanzania_mixed_group_leader][:C]    = s.cell('Q', 77)
    ggl[:tanzania_mixed_group_leader][:D]    = s.cell('Q', 79)
    ggl[:total_female_group_leader][:A]      = s.cell('R', 73)
    ggl[:total_female_group_leader][:B]      = s.cell('R', 75)
    ggl[:total_female_group_leader][:C]      = s.cell('R', 77)
    ggl[:total_female_group_leader][:D]      = s.cell('R', 79)
    ggl[:total_male_group_leader][:A]        = s.cell('S', 73)
    ggl[:total_male_group_leader][:B]        = s.cell('S', 75)
    ggl[:total_male_group_leader][:C]        = s.cell('S', 77)
    ggl[:total_male_group_leader][:D]        = s.cell('S', 79)
    ggl[:total_mixed_group_leader][:A]       = s.cell('T', 73)
    ggl[:total_mixed_group_leader][:B]       = s.cell('T', 75)
    ggl[:total_mixed_group_leader][:C]       = s.cell('T', 77)
    ggl[:total_mixed_group_leader][:D]       = s.cell('T', 79)


    gpgl[:bangladesh_female_group_leader][:A] = s.cell('C', 74)
    gpgl[:bangladesh_female_group_leader][:B] = s.cell('C', 76)
    gpgl[:bangladesh_female_group_leader][:C] = s.cell('C', 78)
    gpgl[:bangladesh_female_group_leader][:D] = s.cell('C', 80)
    gpgl[:bangladesh_male_group_leader][:A]   = s.cell('D', 74)
    gpgl[:bangladesh_male_group_leader][:B]   = s.cell('D', 76)
    gpgl[:bangladesh_male_group_leader][:C]   = s.cell('D', 78)
    gpgl[:bangladesh_male_group_leader][:D]   = s.cell('D', 80)
    gpgl[:bangladesh_mixed_group_leader][:A]  = s.cell('E', 74)
    gpgl[:bangladesh_mixed_group_leader][:B]  = s.cell('E', 76)
    gpgl[:bangladesh_mixed_group_leader][:C]  = s.cell('E', 78)
    gpgl[:bangladesh_mixed_group_leader][:D]  = s.cell('E', 80)
    gpgl[:ghana_female_group_leader][:A]      = s.cell('F', 74)
    gpgl[:ghana_female_group_leader][:B]      = s.cell('F', 76)
    gpgl[:ghana_female_group_leader][:C]      = s.cell('F', 78)
    gpgl[:ghana_female_group_leader][:D]      = s.cell('F', 80)
    gpgl[:ghana_male_group_leader][:A]        = s.cell('G', 74)
    gpgl[:ghana_male_group_leader][:B]        = s.cell('G', 76)
    gpgl[:ghana_male_group_leader][:C]        = s.cell('G', 78)
    gpgl[:ghana_male_group_leader][:D]        = s.cell('G', 80)
    gpgl[:ghana_mixed_group_leader][:A]       = s.cell('H', 74)
    gpgl[:ghana_mixed_group_leader][:B]       = s.cell('H', 76)
    gpgl[:ghana_mixed_group_leader][:C]       = s.cell('H', 78)
    gpgl[:ghana_mixed_group_leader][:D]       = s.cell('H', 80)
    gpgl[:india_female_group_leader][:A]      = s.cell('I', 74)
    gpgl[:india_female_group_leader][:B]      = s.cell('I', 76)
    gpgl[:india_female_group_leader][:C]      = s.cell('I', 78)
    gpgl[:india_female_group_leader][:D]      = s.cell('I', 80)
    gpgl[:india_male_group_leader][:A]        = s.cell('J', 74)
    gpgl[:india_male_group_leader][:B]        = s.cell('J', 76)
    gpgl[:india_male_group_leader][:C]        = s.cell('J', 78)
    gpgl[:india_male_group_leader][:D]        = s.cell('J', 80)
    gpgl[:india_mixed_group_leader][:A]       = s.cell('K', 74)
    gpgl[:india_mixed_group_leader][:B]       = s.cell('K', 76)
    gpgl[:india_mixed_group_leader][:C]       = s.cell('K', 78)
    gpgl[:india_mixed_group_leader][:D]       = s.cell('K', 80)
    gpgl[:malawi_female_group_leader][:A]     = s.cell('L', 74)
    gpgl[:malawi_female_group_leader][:B]     = s.cell('L', 76)
    gpgl[:malawi_female_group_leader][:C]     = s.cell('L', 78)
    gpgl[:malawi_female_group_leader][:D]     = s.cell('L', 80)
    gpgl[:malawi_male_group_leader][:A]       = s.cell('M', 74)
    gpgl[:malawi_male_group_leader][:B]       = s.cell('M', 76)
    gpgl[:malawi_male_group_leader][:C]       = s.cell('M', 78)
    gpgl[:malawi_male_group_leader][:D]       = s.cell('M', 80)
    gpgl[:malawi_mixed_group_leader][:A]      = s.cell('N', 74)
    gpgl[:malawi_mixed_group_leader][:B]      = s.cell('N', 76)
    gpgl[:malawi_mixed_group_leader][:C]      = s.cell('N', 78)
    gpgl[:malawi_mixed_group_leader][:D]      = s.cell('N', 80)
    gpgl[:tanzania_female_group_leader][:A]   = s.cell('O', 74)
    gpgl[:tanzania_female_group_leader][:B]   = s.cell('O', 76)
    gpgl[:tanzania_female_group_leader][:C]   = s.cell('O', 78)
    gpgl[:tanzania_female_group_leader][:D]   = s.cell('O', 80)
    gpgl[:tanzania_male_group_leader][:A]     = s.cell('P', 74)
    gpgl[:tanzania_male_group_leader][:B]     = s.cell('P', 76)
    gpgl[:tanzania_male_group_leader][:C]     = s.cell('P', 78)
    gpgl[:tanzania_male_group_leader][:D]     = s.cell('P', 80)
    gpgl[:tanzania_mixed_group_leader][:A]    = s.cell('Q', 74)
    gpgl[:tanzania_mixed_group_leader][:B]    = s.cell('Q', 76)
    gpgl[:tanzania_mixed_group_leader][:C]    = s.cell('Q', 78)
    gpgl[:tanzania_mixed_group_leader][:D]    = s.cell('Q', 80)
    gpgl[:total_female_group_leader][:A]      = s.cell('R', 74)
    gpgl[:total_female_group_leader][:B]      = s.cell('R', 76)
    gpgl[:total_female_group_leader][:C]      = s.cell('R', 78)
    gpgl[:total_female_group_leader][:D]      = s.cell('R', 80)
    gpgl[:total_male_group_leader][:A]        = s.cell('S', 74)
    gpgl[:total_male_group_leader][:B]        = s.cell('S', 76)
    gpgl[:total_male_group_leader][:C]        = s.cell('S', 78)
    gpgl[:total_male_group_leader][:D]        = s.cell('S', 80)
    gpgl[:total_mixed_group_leader][:A]       = s.cell('T', 74)
    gpgl[:total_mixed_group_leader][:B]       = s.cell('T', 76)
    gpgl[:total_mixed_group_leader][:C]       = s.cell('T', 78)
    gpgl[:total_mixed_group_leader][:D]       = s.cell('T', 80)

    
    bgl[:bangladesh_female_group_leader][:No]  = s.cell('C', 86)
    bgl[:bangladesh_female_group_leader][:Yes] = s.cell('C', 88)
    bgl[:bangladesh_male_group_leader][:No]    = s.cell('D', 86)
    bgl[:bangladesh_male_group_leader][:Yes]   = s.cell('D', 88)
    bgl[:bangladesh_mixed_group_leader][:No]   = s.cell('E', 86)
    bgl[:bangladesh_mixed_group_leader][:Yes]  = s.cell('E', 88)
    bgl[:ghana_female_group_leader][:No]       = s.cell('F', 86)
    bgl[:ghana_female_group_leader][:Yes]      = s.cell('F', 88)
    bgl[:ghana_male_group_leader][:No]         = s.cell('G', 86)
    bgl[:ghana_male_group_leader][:Yes]        = s.cell('G', 88)
    bgl[:ghana_mixed_group_leader][:No]        = s.cell('H', 86)
    bgl[:ghana_mixed_group_leader][:Yes]       = s.cell('H', 88)
    bgl[:india_female_group_leader][:No]       = s.cell('I', 86)
    bgl[:india_female_group_leader][:Yes]      = s.cell('I', 88)
    bgl[:india_male_group_leader][:No]         = s.cell('J', 86)
    bgl[:india_male_group_leader][:Yes]        = s.cell('J', 88)
    bgl[:india_mixed_group_leader][:No]        = s.cell('K', 86)
    bgl[:india_mixed_group_leader][:Yes]       = s.cell('K', 88)
    bgl[:malawi_female_group_leader][:No]      = s.cell('L', 86)
    bgl[:malawi_female_group_leader][:Yes]     = s.cell('L', 88)
    bgl[:malawi_male_group_leader][:No]        = s.cell('M', 86)
    bgl[:malawi_male_group_leader][:Yes]       = s.cell('M', 88)
    bgl[:malawi_mixed_group_leader][:No]       = s.cell('N', 86)
    bgl[:malawi_mixed_group_leader][:Yes]      = s.cell('N', 88)
    bgl[:tanzania_female_group_leader][:No]    = s.cell('O', 86)
    bgl[:tanzania_female_group_leader][:Yes]   = s.cell('O', 88)
    bgl[:tanzania_male_group_leader][:No]      = s.cell('P', 86)
    bgl[:tanzania_male_group_leader][:Yes]     = s.cell('P', 88)
    bgl[:tanzania_mixed_group_leader][:No]     = s.cell('Q', 86)
    bgl[:tanzania_mixed_group_leader][:Yes]    = s.cell('Q', 88)
    bgl[:total_female_group_leader][:No]       = s.cell('R', 86)
    bgl[:total_female_group_leader][:Yes]      = s.cell('R', 88)
    bgl[:total_male_group_leader][:No]         = s.cell('S', 86)
    bgl[:total_male_group_leader][:Yes]        = s.cell('S', 88)
    bgl[:total_mixed_group_leader][:No]        = s.cell('T', 86)
    bgl[:total_mixed_group_leader][:Yes]       = s.cell('T', 88)

    
    bpgl[:bangladesh_female_group_leader][:No]  = s.cell('C', 87)
    bpgl[:bangladesh_female_group_leader][:Yes] = s.cell('C', 89)
    bpgl[:bangladesh_male_group_leader][:No]    = s.cell('D', 87)
    bpgl[:bangladesh_male_group_leader][:Yes]   = s.cell('D', 89)
    bpgl[:bangladesh_mixed_group_leader][:No]   = s.cell('E', 87)
    bpgl[:bangladesh_mixed_group_leader][:Yes]  = s.cell('E', 89)
    bpgl[:ghana_female_group_leader][:No]       = s.cell('F', 87)
    bpgl[:ghana_female_group_leader][:Yes]      = s.cell('F', 89)
    bpgl[:ghana_male_group_leader][:No]         = s.cell('G', 87)
    bpgl[:ghana_male_group_leader][:Yes]        = s.cell('G', 89)
    bpgl[:ghana_mixed_group_leader][:No]        = s.cell('H', 87)
    bpgl[:ghana_mixed_group_leader][:Yes]       = s.cell('H', 89)
    bpgl[:india_female_group_leader][:No]       = s.cell('I', 87)
    bpgl[:india_female_group_leader][:Yes]      = s.cell('I', 89)
    bpgl[:india_male_group_leader][:No]         = s.cell('J', 87)
    bpgl[:india_male_group_leader][:Yes]        = s.cell('J', 89)
    bpgl[:india_mixed_group_leader][:No]        = s.cell('K', 87)
    bpgl[:india_mixed_group_leader][:Yes]       = s.cell('K', 89)
    bpgl[:malawi_female_group_leader][:No]      = s.cell('L', 87)
    bpgl[:malawi_female_group_leader][:Yes]     = s.cell('L', 89)
    bpgl[:malawi_male_group_leader][:No]        = s.cell('M', 87)
    bpgl[:malawi_male_group_leader][:Yes]       = s.cell('M', 89)
    bpgl[:malawi_mixed_group_leader][:No]       = s.cell('N', 87)
    bpgl[:malawi_mixed_group_leader][:Yes]      = s.cell('N', 89)
    bpgl[:tanzania_female_group_leader][:No]    = s.cell('O', 87)
    bpgl[:tanzania_female_group_leader][:Yes]   = s.cell('O', 89)
    bpgl[:tanzania_male_group_leader][:No]      = s.cell('P', 87)
    bpgl[:tanzania_male_group_leader][:Yes]     = s.cell('P', 89)
    bpgl[:tanzania_mixed_group_leader][:No]     = s.cell('Q', 87)
    bpgl[:tanzania_mixed_group_leader][:Yes]    = s.cell('Q', 89)
    bpgl[:total_female_group_leader][:No]       = s.cell('R', 87)
    bpgl[:total_female_group_leader][:Yes]      = s.cell('R', 89)
    bpgl[:total_male_group_leader][:No]         = s.cell('S', 87)
    bpgl[:total_male_group_leader][:Yes]        = s.cell('S', 89)
    bpgl[:total_mixed_group_leader][:No]        = s.cell('T', 87)
    bpgl[:total_mixed_group_leader][:Yes]       = s.cell('T', 89)

   end

