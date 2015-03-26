# Roo::Xls [![BuildStatus](https://travis-ci.org/roo-rb/roo-xls.svg)](https://travis-ci.org/roo-rb/roo-xls)[![Code Climate](https://codeclimate.com/github/roo-rb/roo-xls/badges/gpa.svg)](https://codeclimate.com/github/roo-rb/roo-xls)[![Coverage Status](https://coveralls.io/repos/roo-rb/roo-xls/badge.svg?branch=master)](https://coveralls.io/r/roo-rb/roo-xls?branch=master)

This library extends Roo to add support for handling class Excel files, including:

* .xls files
* .xml files in the SpreadsheetML format (circa 2003)

There is no support for formulas in Roo for .xls files - you can get the result
of a formula but not the formula itself.

## License

While Roo and Roo::Xls are licensed under the MIT / Expat license, please note that the 'spreadsheet' gem [is released under](https://github.com/zdavatz/spreadsheet/blob/master/LICENSE.txt) the GPLv3 license.

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

TODO: Write usage instructions here

## Contributing

1. Fork it ( https://github.com/[my-github-username]/roo-xls/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
