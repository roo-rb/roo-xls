source 'https://rubygems.org'

# Specify your gem's dependencies in roo-xls.gemspec
gemspec

if ENV['CI']
  gem 'roo', '>= 2.0.0beta1', git: 'https://github.com/roo-rb/roo.git'
else
  gem 'roo', '>= 2.0.0beta1', '< 3'
end

group :test do
  # additional testing libs
  gem 'webmock'
  gem 'shoulda'
  gem 'rspec', '>= 3.0.0'
  gem 'simplecov', '>= 0.21', require: false
  gem 'simplecov-lcov', '>= 0.8', require: false
  gem 'coveralls', require: false
end

group :local_development do
  gem 'terminal-notifier-guard', require: false if RUBY_PLATFORM.downcase.include?('darwin')
  gem 'guard-rspec', '>= 4.3.1' ,require: false
  gem 'guard-bundler', require: false
  gem 'guard-preek', require: false
  gem 'guard-rubocop', require: false
  gem 'guard-reek', git: 'https://github.com/pericles/guard-reek', require: false
  gem 'pry'
end
