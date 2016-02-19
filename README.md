# Roo::Xls

[![Build Status](https://img.shields.io/travis/roo-rb/roo-xls.svg?style=flat-square)](https://travis-ci.org/roo-rb/roo-xls) [![Code Climate](https://img.shields.io/codeclimate/github/roo-rb/roo-xls.svg?style=flat-square)](https://codeclimate.com/github/roo-rb/roo-xls) [![Coverage Status](https://img.shields.io/coveralls/roo-rb/roo-xls.svg?style=flat-square)](https://coveralls.io/r/roo-rb/roo-xls) [![Gem Version](https://img.shields.io/gem/v/roo-xls.svg?style=flat-square)](https://rubygems.org/gems/roo-xls)

This library extends Roo to add support for handling class Excel files, including:

* .xls files
* .xml files in the SpreadsheetML format (circa 2003)

There is no support for formulas in Roo for .xls files - you can get the result
of a formula but not the formula itself.

## License

While Roo and Roo::Xls are licensed under the MIT / Expat license, please note that the 'spreadsheet' gem is in a somewhat ambiguous state in terms of licensing.

It is ostensibly [licensed under the GPLv3](https://github.com/zdavatz/spreadsheet/blob/master/LICENSE.txt), but the author [claims you need a commercial license](http://spreadsheet.ch/2014/10/24/using-ruby-spreadsheet-on-heroku-with-dynos/) to use it as part of a public-facing, closed-source service.

This interpretation [is at odds with the FSF's intent and interpretation of the license](http://www.gnu.org/licenses/gpl-faq.html#UnreleasedMods), as this scenario is the reason for the existence of the [GNU AGPLv3](https://www.gnu.org/licenses/agpl.txt).  The maintainer of the 'spreadsheet' gem, however, [disagrees](https://github.com/zdavatz/spreadsheet/issues/167).

As such, until this matter is resolved one way or another, you may wish to take the precautionary step of assuming that use of the 'spreadsheet' gem confers the same obligations as those under the GNU AGPLv3 -- unless a commercial license is obtained -- rather than those of the GPLv3.

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

1. Fork it ( https://github.com/roo-rb/roo-xls/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request
