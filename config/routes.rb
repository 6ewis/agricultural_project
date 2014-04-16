Rails.application.routes.draw do
  root to: 'welcome#index'
  get 'country' => 'welcome#country'
  get 'valuechain' => 'welcome#valuechain'
  get 'group_composition' => 'welcome#group_composition'
  get 'group_leader' => 'welcome#group_leader'
end
