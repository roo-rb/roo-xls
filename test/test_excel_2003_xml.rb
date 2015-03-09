require 'test_helper'

class TestRooExcel < MiniTest::Test
  def with_xml_spreadsheet(name)
    yield ::Roo::Excel2003XML.new(File.join(TESTDIR, "#{name}.xml"))
  end

  def test_xml_namespace_ss
    with_xml_spreadsheet('excel2003') do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal 'BST Variables', oo.cell(1, 1)
    end
  end

  def test_xml_namespace_non_ss
    with_xml_spreadsheet('excel2003_namespace') do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal 'DYS393', oo.cell(1, 1)
      assert_equal '13', oo.cell(2, 1)
    end
  end

  # If a cell has a date-like string but is preceeded by a '
  # to force that date to be treated like a string, we were getting an exception.
  # This test just checks for that exception to make sure it's not raised in this case
  def test_date_to_float_conversion
    with_xml_spreadsheet('datetime_floatconv') do |oo|
      assert_nothing_raised(NoMethodError) do
        oo.cell('a', 1)
        oo.cell('a', 2)
      end
    end
  end

  def test_ruby_spreadsheet_formula_bug
    with_xml_spreadsheet('formula_parse_error') do |oo|
      assert_equal '5026', oo.cell(2, 3)
      assert_equal '5026', oo.cell(3, 3)
    end
  end
end
