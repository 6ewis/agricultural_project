Rails.application.routes.draw do
  root to: 'welcome#index'
  get 'country' => 'welcome#country'
  get 'valuechain' => 'welcome#valuechain'
  get 'gender_composition' => 'welcome#gender_composition'
  get 'gender_leader' => 'welcome#gender_leader'
end
