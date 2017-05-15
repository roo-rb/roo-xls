# -- encoding : utf-8 --
require 'test_helper'
require 'stringio'

class TestRooExcel < MiniTest::Test
  def with_spreadsheet(name)
    yield ::Roo::Spreadsheet.open(File.join(TESTDIR, "#{name}.xls"))
  end

  # Excel can only read the cell's value
  def test_formula_excel
    with_spreadsheet('formula') do |oo|
      assert_equal 21, oo.cell('A', 7)
      assert_equal 21, oo.cell('B', 7)
    end
  end

  # Ruby-spreadsheet now allows us to at least give the current value
  # from a cell with a formula (no possible with parseexcel)
  def test_bug_false_borders_with_formulas
    with_spreadsheet('false_encoding') do |oo|
      assert_equal 1, oo.first_row
      assert_equal 3, oo.last_row
      assert_equal 1, oo.first_column
      assert_equal 4, oo.last_column
    end
  end

  # We'ce added minimal formula support so we can now read these
  # though not sure how the spreadsheet reports older values....
  def test_fe
    with_spreadsheet('false_encoding') do |oo|
      assert_equal Date.new(2007, 11, 1), oo.cell('a', 1)
      # DOES NOT WORK IN EXCEL FILES: assert_equal true, oo.formula?('a',1)
      # DOES NOT WORK IN EXCEL FILES: assert_equal '=TODAY()', oo.formula('a',1)

      assert_equal Date.new(2008, 2, 9), oo.cell('B', 1)
      # DOES NOT WORK IN EXCEL FILES: assert_equal true,               oo.formula?('B',1)
      # DOES NOT WORK IN EXCEL FILES: assert_equal "=A1+100",          oo.formula('B',1)

      assert_kind_of DateTime, oo.cell('C', 1)
      # DOES NOT WORK IN EXCEL FILES: assert_equal true,               oo.formula?('C',1)
      # DOES NOT WORK IN EXCEL FILES: assert_equal "=C1",          oo.formula('C',1)

      assert_equal 'H1', oo.cell('A', 2)
      assert_equal 'H2', oo.cell('B', 2)
      assert_equal 'H3', oo.cell('C', 2)
      assert_equal 'H4', oo.cell('D', 2)
      assert_equal 'R1', oo.cell('A', 3)
      assert_equal 'R2', oo.cell('B', 3)
      assert_equal 'R3', oo.cell('C', 3)
      assert_equal 'R4', oo.cell('D', 3)
    end
  end

  def test_excel_does_not_support_formulas
    with_spreadsheet('false_encoding') do |oo|
      assert_raises(NotImplementedError) { oo.formula('a', 1) }
      assert_raises(NotImplementedError) { oo.formula?('a', 1) }
      assert_raises(NotImplementedError) { oo.formulas(oo.sheets.first) }
    end
  end

  def test_bug_excel_numbers1_sheet5_last_row
    with_spreadsheet('numbers1') do |oo|
      oo.default_sheet = 'Tabelle1'
      assert_equal 1, oo.first_row
      assert_equal 18, oo.last_row
      assert_equal ::Roo::Utils.letter_to_number('A'), oo.first_column
      assert_equal ::Roo::Utils.letter_to_number('G'), oo.last_column
      oo.default_sheet = 'Name of Sheet 2'
      assert_equal 5, oo.first_row
      assert_equal 14, oo.last_row
      assert_equal ::Roo::Utils.letter_to_number('B'), oo.first_column
      assert_equal ::Roo::Utils.letter_to_number('E'), oo.last_column
      oo.default_sheet = 'Sheet3'
      assert_equal 1, oo.first_row
      assert_equal 1, oo.last_row
      assert_equal ::Roo::Utils.letter_to_number('A'), oo.first_column
      assert_equal ::Roo::Utils.letter_to_number('BA'), oo.last_column
      oo.default_sheet = 'Sheet4'
      assert_equal 1, oo.first_row
      assert_equal 1, oo.last_row
      assert_equal ::Roo::Utils.letter_to_number('A'), oo.first_column
      assert_equal ::Roo::Utils.letter_to_number('E'), oo.last_column
      oo.default_sheet = 'Sheet5'
      assert_equal 1, oo.first_row
      assert_equal 6, oo.last_row
      assert_equal ::Roo::Utils.letter_to_number('A'), oo.first_column
      assert_equal ::Roo::Utils.letter_to_number('E'), oo.last_column
    end
  end

  def test_bug_row_column_fixnum_float
    with_spreadsheet('bug-row-column-fixnum-float') do |oo|
      assert_equal 42.5, oo.cell('b', 2)
      assert_equal 43, oo.cell('c', 2)
      assert_equal ['hij', 42.5, 43], oo.row(2)
      assert_equal ['def', 42.5, 'nop'], oo.column(2)
    end
  end

  def test_file_warning_ignore
    Roo::Excel.new(File.join(TESTDIR, 'type_excel.ods'),
                   packed: false,
                   file_warning: :ignore)
    Roo::Excel.new(File.join(TESTDIR, 'type_excel.xlsx'),
                   packed: false,
                   file_warning: :ignore)
  end

  def test_bug_last_row_excel
    with_spreadsheet('time-test') do |oo|
      assert_equal 2, oo.last_row
    end
  end

  def test_excel_download_uri_and_zipped
    if ONLINE
      url = 'http://stiny-leonhard.de/bode-v1.xls.zip'
      excel = Roo::Excel.new(url, packed: :zip)
      excel.default_sheet = excel.sheets.first
      assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
    end
  end

  def test_excel_zipped
    oo = Roo::Excel.new(File.join(TESTDIR, 'bode-v1.xls.zip'), packed: :zip)
    assert oo
    assert_equal 'ist "e" im Nenner von H(s)', oo.cell('b', 5)
  end

  def test_should_raise_file_not_found_error
    assert_raises(IOError) do
      Roo::Excel.new(File.join('testnichtvorhanden', 'Bibelbund.xls'))
    end
  end

  def test_file_warning_default
    assert_raises(TypeError) { Roo::Excel.new(File.join(TESTDIR, 'numbers1.ods')) }
    assert_raises(TypeError) { Roo::Excel.new(File.join(TESTDIR, 'numbers1.xlsx')) }
  end

  def test_file_warning_error
    assert_raises(TypeError) do
      Roo::Excel.new(File.join(TESTDIR, 'numbers1.ods'),
                     packed: false,
                     file_warning: :error)
    end
    assert_raises(TypeError) do
      Roo::Excel.new(File.join(TESTDIR, 'numbers1.xlsx'),
                     packed: false,
                     file_warning: :error)
    end
  end

  def test_file_warning_warning
    assert_raises(Ole::Storage::FormatError) do
      Roo::Excel.new(File.join(TESTDIR, 'numbers1.ods'),
                     packed: false,
                     file_warning: :warning)
    end
    assert_raises(Ole::Storage::FormatError) do
      Roo::Excel.new(File.join(TESTDIR, 'numbers1.xlsx'),
                     packed: false,
                     file_warning: :warning)
    end
  end

  def test_download_uri
    if ONLINE
      assert_raises(RuntimeError) do
        Roo::Excel.new('http://gibbsnichtdomainxxxxx.com/file.xls')
      end
    end
  end

  def test_download_uri_with_query_string
    dir = File.expand_path("#{File.dirname __FILE__}/files")
    file = "#{dir}/simple_spreadsheet.xls"
    url = 'http://test.example.com/simple_spreadsheet.xls?query-param=value'
    stub_request(:any, url).to_return(body: File.read(file))
    spreadsheet = Roo::Excel.new(url)
    spreadsheet.default_sheet = spreadsheet.sheets.first
    assert_equal 'Task 1', spreadsheet.cell('f', 4)
  end

  def test_italo_table
    with_spreadsheet('simple_spreadsheet_from_italo') do |oo|
      assert_equal '1', oo.cell('A', 1)
      assert_equal '1', oo.cell('B', 1)
      assert_equal '1', oo.cell('C', 1)
      assert_equal 1, oo.cell('A', 2).to_i
      assert_equal 2, oo.cell('B', 2).to_i
      assert_equal 1, oo.cell('C', 2).to_i
      assert_equal 1, oo.cell('A', 3)
      assert_equal 3, oo.cell('B', 3)
      assert_equal 1, oo.cell('C', 3)
      assert_equal 'A', oo.cell('A', 4)
      assert_equal 'A', oo.cell('B', 4)
      assert_equal 'A', oo.cell('C', 4)
      assert_equal 0.01, oo.cell('A', 5)
      assert_equal 0.01, oo.cell('B', 5)
      assert_equal 0.01, oo.cell('C', 5)
      assert_equal 0.03, oo.cell('a', 5) + oo.cell('b', 5) + oo.cell('c', 5)

      # Cells values in row 1:
      assert_equal '1:string', oo.cell(1, 1) + ':' + oo.celltype(1, 1).to_s
      assert_equal '1:string', oo.cell(1, 2) + ':' + oo.celltype(1, 2).to_s
      assert_equal '1:string', oo.cell(1, 3) + ':' + oo.celltype(1, 3).to_s

      # Cells values in row 2:
      assert_equal '1:string', oo.cell(2, 1) + ':' + oo.celltype(2, 1).to_s
      assert_equal '2:string', oo.cell(2, 2) + ':' + oo.celltype(2, 2).to_s
      assert_equal '1:string', oo.cell(2, 3) + ':' + oo.celltype(2, 3).to_s

      # Cells values in row 3:
      assert_equal '1.0:float', oo.cell(3, 1).to_s + ':' + oo.celltype(3, 1).to_s
      assert_equal '3.0:float', oo.cell(3, 2).to_s + ':' + oo.celltype(3, 2).to_s
      assert_equal '1.0:float', oo.cell(3, 3).to_s + ':' + oo.celltype(3, 3).to_s

      # Cells values in row 4:
      assert_equal 'A:string', oo.cell(4, 1) + ':' + oo.celltype(4, 1).to_s
      assert_equal 'A:string', oo.cell(4, 2) + ':' + oo.celltype(4, 2).to_s
      assert_equal 'A:string', oo.cell(4, 3) + ':' + oo.celltype(4, 3).to_s

      # Cells values in row 5:
      assert_equal '0.01:float', oo.cell(5, 1).to_s + ':' + oo.celltype(5, 1).to_s
      assert_equal '0.01:float', oo.cell(5, 2).to_s + ':' + oo.celltype(5, 2).to_s
      assert_equal '0.01:float', oo.cell(5, 3).to_s + ':' + oo.celltype(5, 3).to_s
    end
  end

  # "/tmp/xxxx" darf man unter Windows nicht verwenden, weil das nicht erkannt
  # wird.
  # Besser: Methode um temporaeres Dir. portabel zu bestimmen
  def test_huge_document_to_csv
    if LONG_RUN
      with_spreadsheet('Bibelbund') do |oo|
        Dir.mktmpdir do |tempdir|
          assert_equal "Tagebuch des Sekret\303\244rs.    Letzte Tagung 15./16.11.75 Schweiz", oo.cell(45, 'A')
          assert_equal "Tagebuch des Sekret\303\244rs.  Nachrichten aus Chile", oo.cell(46, 'A')
          assert_equal 'Tagebuch aus Chile  Juli 1977', oo.cell(55, 'A')
          assert oo.to_csv(File.join(tempdir, 'Bibelbund.csv'))
          assert File.exist?(File.join(tempdir, 'Bibelbund.csv'))
          assert_equal '', file_diff(File.join(TESTDIR, 'Bibelbund.csv'), File.join(tempdir, 'Bibelbund.csv'))
        end
      end
    end
  end

  def test_bug_quotes_excelx
    if LONG_RUN
      with_spreadsheet('Bibelbund') do |oo|
        oo.default_sheet = oo.sheets.first
        assert_equal 'Einflüsse der neuen Theologie in "de gereformeerde Kerken van Nederland"',
                     oo.cell('a', 76)
        oo.to_csv("csv#{$PROCESS_ID}")
        assert_equal 'Einflüsse der neuen Theologie in "de gereformeerde Kerken van Nederland"',
                     oo.cell('a', 78)
        File.delete_if_exist("csv#{$PROCESS_ID}")
      end
    end
  end

  def test_find_by_row_huge_document
    if LONG_RUN
      with_spreadsheet('Bibelbund') do |oo|
        oo.default_sheet = oo.sheets.first
        rec = oo.find 20
        assert rec
        # assert_equal "Brief aus dem Sekretariat", rec[0]
        # p rec
        assert_equal 'Brief aus dem Sekretariat', rec[0]['TITEL']
        rec = oo.find 22
        assert rec
        # assert_equal "Brief aus dem Skretariat. Tagung in Amberg/Opf.",rec[0]
        assert_equal 'Brief aus dem Skretariat. Tagung in Amberg/Opf.', rec[0]['TITEL']
      end
    end
  end

  def test_find_by_row
    with_spreadsheet('numbers1') do |oo|
      oo.header_line = nil
      rec = oo.find 16
      assert rec
      assert_nil oo.header_line
      # keine Headerlines in diesem Beispiel definiert
      assert_equal 'einundvierzig', rec[0]
      # assert_equal false, rec
      rec = oo.find 15
      assert rec
      assert_equal 41, rec[0]
    end
  end

  def test_find_by_row_if_header_line_is_not_nil
    with_spreadsheet('numbers1') do |oo|
      oo.header_line = 2
      refute_nil oo.header_line
      rec = oo.find 1
      assert rec
      assert_equal 5, rec[0]
      assert_equal 6, rec[1]
      rec = oo.find 15
      assert rec
      assert_equal 'einundvierzig', rec[0]
    end
  end

  def test_find_by_conditions
    if LONG_RUN
      with_spreadsheet('Bibelbund') do |oo|
        #-----------------------------------------------------------------
        zeilen = oo.find(:all, conditions: {
                           'TITEL' => 'Brief aus dem Sekretariat'
                         }
        )
        assert_equal 2, zeilen.size
        assert_equal [{ 'VERFASSER' => 'Almassy, Annelene von',
                        'INTERNET' => nil,
                        'SEITE' => 316.0,
                        'KENNUNG' => 'Aus dem Bibelbund',
                        'OBJEKT' => 'Bibel+Gem',
                        'PC' => '#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#',
                        'NUMMER' => '1982-3',
                        'TITEL' => 'Brief aus dem Sekretariat' },
                      { 'VERFASSER' => 'Almassy, Annelene von',
                        'INTERNET' => nil,
                        'SEITE' => 222.0,
                        'KENNUNG' => 'Aus dem Bibelbund',
                        'OBJEKT' => 'Bibel+Gem',
                        'PC' => '#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#',
                        'NUMMER' => '1983-2',
                        'TITEL' => 'Brief aus dem Sekretariat' }], zeilen

        #----------------------------------------------------------
        zeilen = oo.find(:all,
                         conditions: { 'VERFASSER' => 'Almassy, Annelene von' }
        )
        assert_equal 13, zeilen.size
        #----------------------------------------------------------
        zeilen = oo.find(:all, conditions: {
                           'TITEL' => 'Brief aus dem Sekretariat',
                           'VERFASSER' => 'Almassy, Annelene von'
                         }
        )
        assert_equal 2, zeilen.size
        assert_equal [{ 'VERFASSER' => 'Almassy, Annelene von',
                        'INTERNET' => nil,
                        'SEITE' => 316.0,
                        'KENNUNG' => 'Aus dem Bibelbund',
                        'OBJEKT' => 'Bibel+Gem',
                        'PC' => '#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#',
                        'NUMMER' => '1982-3',
                        'TITEL' => 'Brief aus dem Sekretariat' },
                      { 'VERFASSER' => 'Almassy, Annelene von',
                        'INTERNET' => nil,
                        'SEITE' => 222.0,
                        'KENNUNG' => 'Aus dem Bibelbund',
                        'OBJEKT' => 'Bibel+Gem',
                        'PC' => '#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#',
                        'NUMMER' => '1983-2',
                        'TITEL' => 'Brief aus dem Sekretariat' }], zeilen

        # Result as an array
        zeilen = oo.find(:all,
                         conditions: {
                           'TITEL' => 'Brief aus dem Sekretariat',
                           'VERFASSER' => 'Almassy, Annelene von'
                         }, array: true)
        assert_equal 2, zeilen.size
        assert_equal [
          [
            'Brief aus dem Sekretariat',
            'Almassy, Annelene von',
            'Bibel+Gem',
            '1982-3',
            316.0,
            nil,
            '#C:\\Bibelbund\\reprint\\BuG1982-3.pdf#',
            'Aus dem Bibelbund'
          ],
          [
            'Brief aus dem Sekretariat',
            'Almassy, Annelene von',
            'Bibel+Gem',
            '1983-2',
            222.0,
            nil,
            '#C:\\Bibelbund\\reprint\\BuG1983-2.pdf#',
            'Aus dem Bibelbund'
          ]], zeilen
      end
    end
  end

  # TODO: temporaerer Test
  def test_seiten_als_date
    if LONG_RUN
      with_spreadsheet('Bibelbund', format: :excelx) do |oo|
        assert_equal 'Bericht aus dem Sekretariat', oo.cell(13, 1)
        assert_equal '1981-4', oo.cell(13, 'D')
        assert_equal String, oo.excelx_type(13, 'E')[1].class
        assert_equal [:numeric_or_formula, 'General'], oo.excelx_type(13, 'E')
        assert_equal '428', oo.excelx_value(13, 'E')
        assert_equal 428.0, oo.cell(13, 'E')
      end
    end
  end

  def test_column
    with_spreadsheet('numbers1') do |oo|
      expected = [1.0, 5.0, nil, 10.0, Date.new(1961, 11, 21), 'tata', nil, nil, nil, nil, 'thisisa11', 41.0, nil, nil, 41.0, 'einundvierzig', nil, Date.new(2007, 5, 31)]
      assert_equal expected, oo.column(1)
      assert_equal expected, oo.column('a')
    end
  end

  def test_column_huge_document
    if LONG_RUN
      with_spreadsheet('Bibelbund') do |oo|
        oo.default_sheet = oo.sheets.first
        assert_equal 3735, oo.column('a').size
        # assert_equal 499, oo.column('a').size
      end
    end
  end

  def test_simple_spreadsheet_find_by_condition
    with_spreadsheet('simple_spreadsheet') do |oo|
      oo.header_line = 3
      # oo.date_format = '%m/%d/%Y' if oo.class == Google
      erg = oo.find(:all, conditions: { 'Comment' => 'Task 1' })
      assert_equal Date.new(2007, 05, 07), erg[1]['Date']
      assert_equal 10.75, erg[1]['Start time']
      assert_equal 12.50, erg[1]['End time']
      assert_equal 0, erg[1]['Pause']
      assert_equal 1.75, erg[1]['Sum']
      assert_equal 'Task 1', erg[1]['Comment']
    end
  end

  def test_info
    expected_templ = "File: numbers1%s\n"\
      "Number of sheets: 5\n"\
      "Sheets: Tabelle1, Name of Sheet 2, Sheet3, Sheet4, Sheet5\n"\
      "Sheet 1:\n"\
      "  First row: 1\n"\
      "  Last row: 18\n"\
      "  First column: A\n"\
      "  Last column: G\n"\
      "Sheet 2:\n"\
      "  First row: 5\n"\
      "  Last row: 14\n"\
      "  First column: B\n"\
      "  Last column: E\n"\
      "Sheet 3:\n"\
      "  First row: 1\n"\
      "  Last row: 1\n"\
      "  First column: A\n"\
      "  Last column: BA\n"\
      "Sheet 4:\n"\
      "  First row: 1\n"\
      "  Last row: 1\n"\
      "  First column: A\n"\
      "  Last column: E\n"\
      "Sheet 5:\n"\
      "  First row: 1\n"\
      "  Last row: 6\n"\
      "  First column: A\n"\
      '  Last column: E'
    with_spreadsheet('numbers1') do |oo|
      expected = sprintf(expected_templ, '.xls')
      begin
        assert_equal expected, oo.info
      rescue NameError
        #
      end
    end
  end

  def test_info_doesnt_set_default_sheet
    with_spreadsheet('numbers1') do |oo|
      oo.default_sheet = 'Sheet3'
      oo.info
      assert_equal 'Sheet3', oo.default_sheet
    end
  end

  def test_bug_bbu
    with_spreadsheet('bbu') do |oo|
      assert_equal "File: bbu.xls
Number of sheets: 3
Sheets: 2007_12, Tabelle2, Tabelle3
Sheet 1:
  First row: 1
  Last row: 4
  First column: A
  Last column: F
Sheet 2:
  - empty -
Sheet 3:
  - empty -", oo.info

      oo.default_sheet = oo.sheets[1] # empty sheet
      assert_nil oo.first_row
      assert_nil oo.last_row
      assert_nil oo.first_column
      assert_nil oo.last_column
    end
  end

  def test_bug_time_nil
    with_spreadsheet('time-test') do |oo|
      assert_equal 12 * 3600 + 13 * 60 + 14, oo.cell('B', 1) # 12:13:14 (secs since midnight)
      assert_equal :time, oo.celltype('B', 1)
      assert_equal 15 * 3600 + 16 * 60, oo.cell('C', 1) # 15:16    (secs since midnight)
      assert_equal :time, oo.celltype('C', 1)
      assert_equal 23 * 3600, oo.cell('D', 1) # 23:00    (secs since midnight)
      assert_equal :time, oo.celltype('D', 1)
    end
  end

  def test_date_time_to_csv
    with_spreadsheet('time-test') do |oo|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir, 'time_test.csv')
        assert oo.to_csv(csv_output)
        assert File.exist?(csv_output)
        assert_equal '', `diff --strip-trailing-cr #{TESTDIR}/time-test.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  def test_boolean_to_csv
    with_spreadsheet('boolean') do |oo|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir, 'boolean.csv')
        assert oo.to_csv(csv_output)
        assert File.exist?(csv_output)
        assert_equal '', `diff --strip-trailing-cr #{TESTDIR}/boolean.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  def test_link_to_csv
    with_spreadsheet('link') do |oo|
      Dir.mktmpdir do |tempdir|
        csv_output = File.join(tempdir, 'link.csv')
        assert oo.to_csv(csv_output)
        assert File.exist?(csv_output)
        assert_equal '', `diff --strip-trailing-cr #{TESTDIR}/link.csv #{csv_output}`
        # --strip-trailing-cr is needed because the test-file use 0A and
        # the test on an windows box generates 0D 0A as line endings
      end
    end
  end

  def test_date_time_yaml
    with_spreadsheet('time-test') do |oo|
      expected =
        "--- \ncell_1_1: \n  row: 1 \n  col: 1 \n  celltype: string \n  value: Mittags: \ncell_1_2: \n  row: 1 \n  col: 2 \n  celltype: time \n  value: 12:13:14 \ncell_1_3: \n  row: 1 \n  col: 3 \n  celltype: time \n  value: 15:16:00 \ncell_1_4: \n  row: 1 \n  col: 4 \n  celltype: time \n  value: 23:00:00 \ncell_2_1: \n  row: 2 \n  col: 1 \n  celltype: date \n  value: 2007-11-21 \n"
      assert_equal expected, oo.to_yaml
    end
  end

  # Erstellt eine Liste aller Zellen im Spreadsheet. Dies ist nötig, weil ein einfacher
  # Textvergleich des XML-Outputs nicht funktioniert, da xml-builder die Attribute
  # nicht immer in der gleichen Reihenfolge erzeugt.
  def init_all_cells(oo, sheet)
    all = []
    oo.first_row(sheet).upto(oo.last_row(sheet)) do |row|
      oo.first_column(sheet).upto(oo.last_column(sheet)) do |col|
        unless oo.empty?(row, col, sheet)
          all << { row: row.to_s,
                   column: col.to_s,
                   content: oo.cell(row, col, sheet).to_s,
                   type: oo.celltype(row, col, sheet).to_s
          }
        end
      end
    end
    all
  end

  def test_to_xml
    with_spreadsheet('numbers1') do |oo|
      oo.to_xml
      sheetname = oo.sheets.first
      doc = Nokogiri::XML(oo.to_xml)
      sheet_count = 0
      doc.xpath('//spreadsheet/sheet').each do|_tmpelem|
        sheet_count += 1
      end
      assert_equal 5, sheet_count
      doc.xpath('//spreadsheet/sheet').each do |xml_sheet|
        all_cells = init_all_cells(oo, sheetname)
        x = 0
        assert_equal sheetname, xml_sheet.attributes['name'].value
        xml_sheet.children.each do|cell|
          if cell.attributes['name']
            expected = [all_cells[x][:row],
                        all_cells[x][:column],
                        all_cells[x][:content],
                        all_cells[x][:type]
                       ]
            result = [
              cell.attributes['row'],
              cell.attributes['column'],
              cell.content,
              cell.attributes['type']
            ]
            assert_equal expected, result
            x += 1
          end # if
        end # end of sheet
        sheetname = oo.sheets[oo.sheets.index(sheetname) + 1]
      end
    end
  end

  def test_bug_to_xml_with_empty_sheets
    with_spreadsheet('emptysheets') do |oo|
      oo.sheets.each do |sheet|
        assert_equal nil, oo.first_row, "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row, "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column, "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column, "last_column not nil in sheet #{sheet}"
        assert_equal nil, oo.first_row(sheet), "first_row not nil in sheet #{sheet}"
        assert_equal nil, oo.last_row(sheet), "last_row not nil in sheet #{sheet}"
        assert_equal nil, oo.first_column(sheet), "first_column not nil in sheet #{sheet}"
        assert_equal nil, oo.last_column(sheet), "last_column not nil in sheet #{sheet}"
      end
      oo.to_xml
    end
  end

  def test_datetime
    with_spreadsheet('datetime') do |oo|
      val = oo.cell('c', 3)
      assert_kind_of DateTime, val
      assert_equal :datetime, oo.celltype('c', 3)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), val
      val = oo.cell('a', 1)
      assert_kind_of Date, val
      assert_equal :date, oo.celltype('a', 1)
      assert_equal Date.new(1961, 11, 21), val
      assert_equal Date.new(1961, 11, 21), oo.cell('a', 1)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('a', 3)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('b', 3)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('c', 3)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('a', 4)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('b', 4)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('c', 4)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('a', 5)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('b', 5)
      assert_equal DateTime.new(1961, 11, 21, 12, 17, 18), oo.cell('c', 5)
      assert_equal Date.new(1961, 11, 21), oo.cell('a', 6)
      assert_equal Date.new(1961, 11, 21), oo.cell('b', 6)
      assert_equal Date.new(1961, 11, 21), oo.cell('c', 6)
      assert_equal Date.new(1961, 11, 21), oo.cell('a', 7)
      assert_equal Date.new(1961, 11, 21), oo.cell('b', 7)
      assert_equal Date.new(1961, 11, 21), oo.cell('c', 7)
      assert_equal DateTime.new(2013, 11, 5, 11, 45, 00), oo.cell('a', 8)
      assert_equal DateTime.new(2013, 11, 5, 11, 45, 00), oo.cell('b', 8)
      assert_equal DateTime.new(2013, 11, 5, 11, 45, 00), oo.cell('c', 8)
    end
  end

  def test_cell_boolean
    with_spreadsheet('boolean') do |oo|
      assert_equal 'true', oo.cell(1, 1)
      assert_equal 'false', oo.cell(2, 1)
    end
  end

  def test_cell_multiline
    with_spreadsheet('paragraph') do |oo|
      assert_equal "This is a test\nof a multiline\nCell", oo.cell(1, 1)
      assert_equal "This is a test\n¶\nof a multiline\n\nCell", oo.cell(1, 2)
      assert_equal "first p\n\nsecond p\n\nlast p", oo.cell(2, 1)
    end
  end

  def test_cell_styles
    # styles only valid in excel spreadsheets?
    # TODO: what todo with other spreadsheet types
    with_spreadsheet('style') do |oo|
      # bold
      assert_equal true,  oo.font(1, 1).bold?
      assert_equal false, oo.font(1, 1).italic?
      assert_equal false, oo.font(1, 1).underline?

      # italic
      assert_equal false, oo.font(2, 1).bold?
      assert_equal true,  oo.font(2, 1).italic?
      assert_equal false, oo.font(2, 1).underline?

      # normal
      assert_equal false, oo.font(3, 1).bold?
      assert_equal false, oo.font(3, 1).italic?
      assert_equal false, oo.font(3, 1).underline?

      # underline
      assert_equal false, oo.font(4, 1).bold?
      assert_equal false, oo.font(4, 1).italic?
      assert_equal true,  oo.font(4, 1).underline?

      # bold italic
      assert_equal true,  oo.font(5, 1).bold?
      assert_equal true,  oo.font(5, 1).italic?
      assert_equal false, oo.font(5, 1).underline?

      # bold underline
      assert_equal true,  oo.font(6, 1).bold?
      assert_equal false, oo.font(6, 1).italic?
      assert_equal true,  oo.font(6, 1).underline?

      # italic underline
      assert_equal false, oo.font(7, 1).bold?
      assert_equal true,  oo.font(7, 1).italic?
      assert_equal true,  oo.font(7, 1).underline?

      # bolded row
      assert_equal true, oo.font(8, 1).bold?
      assert_equal false,  oo.font(8, 1).italic?
      assert_equal false,  oo.font(8, 1).underline?

      # bolded col
      assert_equal true, oo.font(9, 2).bold?
      assert_equal false,  oo.font(9, 2).italic?
      assert_equal false,  oo.font(9, 2).underline?

      # bolded row, italic col
      assert_equal true, oo.font(10, 3).bold?
      assert_equal true,  oo.font(10, 3).italic?
      assert_equal false,  oo.font(10, 3).underline?

      # normal
      assert_equal false, oo.font(11, 4).bold?
      assert_equal false,  oo.font(11, 4).italic?
      assert_equal false,  oo.font(11, 4).underline?
    end
  end

  # If a cell has a date-like string but is preceeded by a '
  # to force that date to be treated like a string, we were getting an exception.
  # This test just checks for that exception to make sure it's not raised in this case
  def test_date_to_float_conversion
    with_spreadsheet('datetime_floatconv') do |oo|
      oo.cell('a', 1)
      oo.cell('a', 2)
    end
  end

  # Need to extend to other formats
  def test_row_whitespace
    # auf dieses Dokument habe ich keinen Zugriff TODO:
    with_spreadsheet('whitespace') do |oo|
      oo.default_sheet = 'Sheet1'
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(1)
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(2)
      assert_equal ['Date', 'Start time', 'End time', 'Pause', 'Sum', 'Comment'], oo.row(3)
      assert_equal [Date.new(2007, 5, 7), 9.25, 10.25, 0.0, 1.0, 'Task 1'], oo.row(4)
      assert_equal [nil, nil, nil, nil, nil, nil], oo.row(5)
      assert_equal [Date.new(2007, 5, 7), 10.75, 10.75, 0.0, 0.0, 'Task 1'], oo.row(6)
      oo.default_sheet = 'Sheet2'
      assert_equal ['Date', nil, 'Start time'], oo.row(1)
      assert_equal [Date.new(2007, 5, 7), nil, 9.25], oo.row(2)
      assert_equal [Date.new(2007, 5, 7), nil,  10.75], oo.row(3)
    end
  end

  def test_col_whitespace
    # TODO:
    # kein Zugriff auf Dokument whitespace
    with_spreadsheet('whitespace') do |oo|
      oo.default_sheet = 'Sheet1'
      assert_equal ['Date', Date.new(2007, 5, 7), nil, Date.new(2007, 5, 7)], oo.column(1)
      assert_equal ['Start time', 9.25, nil, 10.75], oo.column(2)
      assert_equal ['End time', 10.25, nil, 10.75], oo.column(3)
      assert_equal ['Pause', 0.0, nil, 0.0], oo.column(4)
      assert_equal ['Sum', 1.0, nil, 0.0], oo.column(5)
      assert_equal ['Comment', 'Task 1', nil, 'Task 1'], oo.column(6)
      oo.default_sheet = 'Sheet2'
      assert_equal [nil, nil, nil], oo.column(1)
      assert_equal [nil, nil, nil], oo.column(2)
      assert_equal ['Date', Date.new(2007, 5, 7), Date.new(2007, 5, 7)], oo.column(3)
      assert_equal [nil, nil, nil], oo.column(4)
      assert_equal ['Start time', 9.25, 10.75], oo.column(5)
    end
  end

  def test_ruby_spreadsheet_formula_bug
    with_spreadsheet('formula_parse_error') do |oo|
      assert_equal '5026', oo.cell(2, 3)
      assert_equal '5026', oo.cell(3, 3)
    end
  end

  def test_excel_links
    with_spreadsheet('link') do |oo|
      assert_equal 'Google', oo.cell(1, 1)
      assert_equal 'http://www.google.com', oo.cell(1, 1).url
    end
  end

  # Excel has two base date formats one from 1900 and the other from 1904.
  # There's a MS bug that 1900 base dates include an extra day due to erroneously
  # including 1900 as a leap yar.
  def test_base_dates
    with_spreadsheet('1900_base') do |oo|
      assert_equal Date.new(2009, 06, 15), oo.cell(1, 1)
      # we don't want to to 'interpret' formulas  assert_equal Date.new(Time.now.year,Time.now.month,Time.now.day), oo.cell(2,1) #formula for TODAY()
      # if we test TODAY() we have also have to calculate
      # other date calculations
      #
      assert_equal :date, oo.celltype(1, 1)
    end
    with_spreadsheet('1904_base') do |oo|
      assert_equal Date.new(2009, 06, 15), oo.cell(1, 1)
      # see comment above
      # assert_equal Date.new(Time.now.year,Time.now.month,Time.now.day), oo.cell(2,1) #formula for TODAY()
      assert_equal :date, oo.celltype(1, 1)
    end
  end

  def test_bad_date
    with_spreadsheet('prova') do |oo|
      assert_equal DateTime.new(2006, 2, 2, 10, 0, 0), oo.cell('a', 1)
    end
  end

  def test_bad_excel_date
    with_spreadsheet('bad_excel_date') do |oo|
      assert_equal DateTime.new(2006, 2, 2, 10, 0, 0), oo.cell('a', 1)
    end
  end

  def test_cell_methods
    with_spreadsheet('numbers1') do |oo|
      assert_equal 10, oo.a4 # cell(4,'A')
      assert_equal 11, oo.b4 # cell(4,'B')
      assert_equal 12, oo.c4 # cell(4,'C')
      assert_equal 13, oo.d4 # cell(4,'D')
      assert_equal 14, oo.e4 # cell(4,'E')
      assert_equal 'ABC', oo.c6('Sheet5')

      # assert_raises(ArgumentError) {
      assert_raises(NoMethodError) do
        # a42a is not a valid cell name, should raise ArgumentError
        assert_equal 9999, oo.a42a
      end
    end
  end

  # compare large spreadsheets
  def test_compare_large_spreadsheets
    # problematisch, weil Formeln in Excel nicht unterstützt werden
    if LONG_RUN
      qq = Roo::OpenOffice.new(File.join('test', 'Bibelbund.ods'))
      with_spreadsheet('Bibelbund') do |oo|
        # p "comparing Bibelbund.ods with #{oo.class}"
        oo.sheets.each do |sh|
          oo.first_row.upto(oo.last_row) do |row|
            oo.first_column.upto(oo.last_column) do |col|
              c1 = qq.cell(row, col, sh)
              c1.force_encoding('UTF-8') if c1.class == String
              c2 = oo.cell(row, col, sh)
              c2.force_encoding('UTF-8') if c2.class == String
              assert_equal c1, c2, "diff in #{sh}/#{row}/#{col}}"
              assert_equal qq.celltype(row, col, sh), oo.celltype(row, col, sh)
            end
          end
        end
      end
    end # LONG_RUN
  end

  require 'matrix'
  def test_matrix
    with_spreadsheet('matrix') do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal Matrix[
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, 9.0]], oo.to_matrix
    end
  end

  def test_matrix_selected_range
    with_spreadsheet('matrix') do |oo|
      oo.default_sheet = 'Sheet2'
      assert_equal Matrix[
        [1.0, 2.0, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, 9.0]], oo.to_matrix(3, 4, 5, 6)
    end
  end

  def test_matrix_all_nil
    with_spreadsheet('matrix') do |oo|
      oo.default_sheet = 'Sheet2'
      assert_equal Matrix[
        [nil, nil, nil],
        [nil, nil, nil],
        [nil, nil, nil]], oo.to_matrix(10, 10, 12, 12)
    end
  end

  def test_matrix_values_and_nil
    with_spreadsheet('matrix') do |oo|
      oo.default_sheet = 'Sheet3'
      assert_equal Matrix[
        [1.0, nil, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, nil]], oo.to_matrix(1, 1, 3, 3)
    end
  end

  def test_matrix_specifying_sheet
    with_spreadsheet('matrix') do |oo|
      oo.default_sheet = oo.sheets.first
      assert_equal Matrix[
        [1.0, nil, 3.0],
        [4.0, 5.0, 6.0],
        [7.0, 8.0, nil]], oo.to_matrix(nil, nil, nil, nil, 'Sheet3')
    end
  end

  # 2011-08-03
  def test_bug_datetime_to_csv
    with_spreadsheet('datetime') do |oo|
      Dir.mktmpdir do |tempdir|
        datetime_csv_file = File.join(tempdir, 'datetime.csv')

        assert oo.to_csv(datetime_csv_file)
        assert File.exist?(datetime_csv_file)
        assert_equal '', file_diff('test/files/so_datetime.csv', datetime_csv_file)
      end
    end
  end

  def common_possible_bug_snowboard_cells(ss)
    assert_equal 'A.', ss.cell(13, 'A'), ss.class
    assert_equal 147, ss.cell(13, 'f'), ss.class
    assert_equal 152, ss.cell(13, 'g'), ss.class
    assert_equal 156, ss.cell(13, 'h'), ss.class
    assert_equal 158, ss.cell(13, 'i'), ss.class
    assert_equal 160, ss.cell(13, 'j'), ss.class
    assert_equal 164, ss.cell(13, 'k'), ss.class
    assert_equal 168, ss.cell(13, 'l'), ss.class
    assert_equal :string, ss.celltype(13, 'm'), ss.class
    assert_equal '159W', ss.cell(13, 'm'), ss.class
    assert_equal '164W', ss.cell(13, 'n'), ss.class
    assert_equal '168W', ss.cell(13, 'o'), ss.class
  end

  # def test_false_encoding
  #   ex = Roo::Excel.new(File.join(TESTDIR,'false_encoding.xls'))
  #   ex.default_sheet = ex.sheets.first
  #   assert_equal "Sheet1", ex.sheets.first
  #   ex.first_row.upto(ex.last_row) do |row|
  #     ex.first_column.upto(ex.last_column) do |col|
  #       content = ex.cell(row,col)
  #       puts "#{row}/#{col}"
  #       #puts content if ! ex.empty?(row,col) or ex.formula?(row,col)
  #       if ex.formula?(row,col)
  #         #! ex.empty?(row,col)
  #         puts content
  #       end
  #     end
  #   end
  # end

  # def test_soap_server
  #   #threads = []
  #   #threads << Thread.new("serverthread") do
  #   fork do
  #     p "serverthread started"
  #     puts "in child, pid = #$$"
  #     puts `/usr/bin/ruby rooserver.rb`
  #     p "serverthread finished"
  #   end
  #   #threads << Thread.new("clientthread") do
  #   p "clientthread started"
  #   sleep 10
  #   proxy = SOAP::RPC::Driver.new("http://localhost:12321","spreadsheetserver")
  #   proxy.add_method('cell','row','col')
  #   proxy.add_method('officeversion')
  #   proxy.add_method('last_row')
  #   proxy.add_method('last_column')
  #   proxy.add_method('first_row')
  #   proxy.add_method('first_column')
  #   proxy.add_method('sheets')
  #   proxy.add_method('set_default_sheet','s')
  #   proxy.add_method('ferien_fuer_region', 'region')

  #   sheets = proxy.sheets
  #   p sheets
  #   proxy.set_default_sheet(sheets.first)

  #   assert_equal 1, proxy.first_row
  #   assert_equal 1, proxy.first_column
  #   assert_equal 187, proxy.last_row
  #   assert_equal 7, proxy.last_column
  #   assert_equal 42, proxy.cell('C',8)
  #   assert_equal 43, proxy.cell('F',12)
  #   assert_equal "1.0", proxy.officeversion
  #   p "clientthread finished"
  #   #end
  #   #threads.each {|t| t.join }
  #   puts "fertig"
  #   Process.kill("INT",pid)
  #   pid = Process.wait
  #   puts "child terminated, pid= #{pid}, status= #{$?.exitstatus}"
  # end

  def split_coord(s)
    letter = ''
    number = 0
    i = 0
    while i < s.length && 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'.include?(s[i, 1])
      letter += s[i, 1]
      i += 1
    end
    while i < s.length && '01234567890'.include?(s[i, 1])
      number = number * 10 + s[i, 1].to_i
      i += 1
    end
    if letter == '' || number == 0
      raise ArgumentError
    end
    [letter, number]
  end

  # def sum(s,expression)
  #  arg = expression.split(':')
  #  b,z = split_coord(arg[0])
  #  first_row = z
  #  first_col = OpenOffice.letter_to_number(b)
  #  b,z = split_coord(arg[1])
  #  last_row = z
  #  last_col = OpenOffice.letter_to_number(b)
  #  result = 0
  #  first_row.upto(last_row) {|row|
  #    first_col.upto(last_col) {|col|
  #      result = result + s.cell(row,col)
  #    }
  #  }
  #  result
  # end

  # def test_create_spreadsheet1
  #  name = File.join(TESTDIR,'createdspreadsheet.ods')
  #  rm(name) if File.exists?(File.join(TESTDIR,'createdspreadsheet.ods'))
  #  # anlegen, falls noch nicht existierend
  #  s = OpenOffice.new(name,true)
  #  assert File.exists?(name)
  # end

  # We don't have the bode-v1.xlsx test file
  # #TODO: xlsx-Datei anpassen!
  # def test_excelx_download_uri_and_zipped
  #   #TODO: gezippte xlsx Datei online zum Testen suchen
  #   if EXCELX
  #     if ONLINE
  #       url = 'http://stiny-leonhard.de/bode-v1.xlsx.zip'
  #       excel = Roo::Excelx.new(url, :zip)
  #       assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
  #     end
  #   end
  # end

  # def test_excelx_zipped
  #   # TODO: bode...xls bei Gelegenheit nach .xlsx konverieren lassen und zippen!
  #   if EXCELX
  #     # diese Datei gibt es noch nicht gezippt
  #     excel = Roo::Excelx.new(File.join(TESTDIR,"bode-v1.xlsx.zip"), :zip)
  #     assert excel
  #     assert_raises(ArgumentError) {
  #       assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
  #     }
  #     excel.default_sheet = excel.sheets.first
  #     assert_equal 'ist "e" im Nenner von H(s)', excel.cell('b', 5)
  #   end
  # end

  def test_excel_via_stringio
    io = StringIO.new(
      File.read(File.join(TESTDIR, 'simple_spreadsheet.xls')))
    spreadsheet = ::Roo::Spreadsheet.open(io, extension: '.xls')
    spreadsheet.default_sheet = spreadsheet.sheets.first
    assert_equal 'Task 1', spreadsheet.cell('f', 4)
  end
end
