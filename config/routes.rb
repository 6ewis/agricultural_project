Rails.application.routes.draw do
  root to: 'welcome#index'
  get 'index' => 'welcome#index'
  get 'country' => 'welcome#country'
  get 'valuechain' => 'welcome#valuechain'
  get 'groupgender' => 'welcome#groupgender'
end
