# Roo::Xls

[![Build Status](https://img.shields.io/travis/roo-rb/roo-xls.svg?style=flat-square)](https://travis-ci.org/roo-rb/roo-xls) [![Code Climate](https://img.shields.io/codeclimate/github/roo-rb/roo-xls.svg?style=flat-square)](https://codeclimate.com/github/roo-rb/roo-xls) [![Coverage Status](https://img.shields.io/coveralls/roo-rb/roo-xls.svg?style=flat-square)](https://coveralls.io/r/roo-rb/roo-xls) [![Gem Version](https://img.shields.io/gem/v/roo-xls.svg?style=flat-square)](https://rubygems.org/gems/roo-xls)

This library extends Roo to add support for handling class Excel files, including:

* .xls files
* .xml files in the SpreadsheetML format (circa 2003)

There is no support for formulas in Roo for .xls files - you can get the result
of a formula but not the formula itself.

## Limitations

Roo::Xls currently doesn't provide support for the following features in Roo:
* [Option `:expand_merged_ranged => true`](https://github.com/roo-rb/roo#expand_merged_ranges)

## License

While Roo and Roo::Xls are licensed under the MIT / Expat license, please note that the `spreadsheet` gem [is released under](https://github.com/zdavatz/spreadsheet/blob/master/LICENSE.txt) the GPLv3 license. Please be aware that the author of the `spreadsheet` gem [claims you need a commercial license](http://spreadsheet.ch/2014/10/24/using-ruby-spreadsheet-on-heroku-with-dynos/) to use it as part of a public-facing, closed-source service, an interpretation [at odds with the FSF's intent and interpretation of the license](http://www.gnu.org/licenses/gpl-faq.html#UnreleasedMods). 

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'roo-xls'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install roo-xls

## Usage

```
begin
  require "bundler/inline"
rescue LoadError => e
  $stderr.puts "Bundler version 1.10 or later is required. Please update your Bundler"
  raise e
end

gemfile(true) do
  source "https://rubygems.org"

  git_source(:github) { |repo| "https://github.com/#{repo}.git" }

  gem "roo-xls"
  gem "minitest"
end

require "roo-xls"
require "minitest/autorun"

class BugTest < Minitest::Test
  def test_stuff
    sheet = Roo::Excel.new('/Users/gturner/downloads/table.xls')
    puts sheet.row(1)
  end
end
```

## Contributing

1. Fork it ( https://github.com/roo-rb/roo-xls/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
