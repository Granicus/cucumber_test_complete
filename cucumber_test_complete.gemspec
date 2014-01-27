Gem::Specification.new do |s|
  s.name        = 'cucumber_test_complete'
  s.version     = '0.1.6'
  s.date        = '2014-01-27'
  s.summary     = 'Test Complete for Cucumber'
  s.description = 'A gem for integrating TestComplete into your Cucumber Ruby World'
  s.authors     = ['Nathan Holland']
  s.email       = 'NathanH@granicus.com'
  s.files       = ['lib/cucumber_test_complete.rb']
  s.homepage    = 'http://rubygems.org/gems/cucumber_test_complete'
  s.license     = 'MIT'
  s.requirements << 'win32ole'
  s.requirements << 'rspec-expectations'
end
