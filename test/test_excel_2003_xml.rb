require 'test_helper'

class TestRooExcel < MiniTest::Test
  def with_spreadsheet(name)
    yield Roo::Spreadsheet.open(File.join(TESTDIR, "#{name}.xml"))
  end

  # If a cell has a date-like string but is preceeded by a '
  # to force that date to be treated like a string, we were getting an exception.
  # This test just checks for that exception to make sure it's not raised in this case
  def test_date_to_float_conversion
    with_each_spreadsheet('datetime_floatconv') do |oo|
      assert_nothing_raised(NoMethodError) do
        oo.cell('a', 1)
        oo.cell('a', 2)
      end
    end
  end

  def test_ruby_spreadsheet_formula_bug
    with_each_spreadsheet('formula_parse_error') do |oo|
      assert_equal '5026', oo.cell(2, 3)
      assert_equal '5026', oo.cell(3, 3)
    end
  end
end
