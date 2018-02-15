require 'spec_helper'

RSpec.describe Roo::Excel2003XML do
  subject { instance }

  let(:instance) { described_class.new(path) }
  let(:path) { File.join(test_files_dir, test_file) }
  let(:test_files_dir) { 'test/files' }

  describe '.new' do
    context 'with an xml file' do
      let(:test_file) { 'datetime.xml' }

      it 'loads the file without error' do
        expect { subject }.to_not raise_error
      end
    end
  end

  describe '#cell' do
    subject { super().cell(cell[:row], cell[:col]) }

    before(:each) { instance.default_sheet = instance.sheets.first }

    context 'with merged cells' do
      # See font_colors_screenshot_in_Mac_Excel_16.10.png for a screenshot of
      # this how this file looks in Mac Excel 16.10.
      let(:test_file) { 'font_colors.xml' }

      context 'the top-left cell in a merged cell' do
        let(:cell) { { :row => 1, :col => 'A' } }

        it 'returns the contents of the merged cell' do
          is_expected.to eq 'Roo::Xls Test of Font Colors'
        end
      end

      context 'the cell to the right of the top-left cell in a merged cell' do
        let(:cell) { { :row => 2, :col => 'A' } }
        it { is_expected.to be_nil }
      end

      context 'the cell below the top-left cell in a merged cell' do
        let(:cell) { { :row => 2, :col => 'B' } }
        it { is_expected.to be_nil }
      end

      context 'the first cell to the right of an entire merged cell' do
        let(:cell) { { :row => 1, :col => 'K' } }

        it 'returns the expected contents' do
          is_expected.to eq 'This entire COLUMN should be ITALIC and GREEN'
        end
      end

      context 'the first cell below an entire merged cell' do
        let(:cell) { { :row => 6, :col => 'A' } }

        it 'returns the expected contents' do
          is_expected.to eq '(The above should be font "Courier New", size 24)'
        end
      end
    end
  end

  describe '#font' do
    subject { super().font(cell[:row], cell[:col]) }

    before(:each) { instance.default_sheet = instance.sheets.first }

    # See font_colors_screenshot_in_Mac_Excel_16.10.png for a screenshot of
    # this how this file looks in Mac Excel 16.10.
    let(:test_file) { 'font_colors.xml' }

    let(:default_attrs) do
      {
        :name       => 'Arial',
        :size       => '12',
        :color      => '#000000',
        :bold?      => false,
        :italic?    => false,
        :underline? => false
      }
    end

    let(:expected_attrs) { default_attrs }

    context 'with no font styling' do
      let(:cell) { { :row => 6, :col => 'A' } }

      it 'returns default font attributes' do
        is_expected.to have_attributes(default_attrs)
      end
    end

    context 'with styling set on an individual cell' do
      context 'when set font name and size' do
        let(:cell) { { :row => 1, :col => 'A' } }

        it 'returns expected font attributes including name and size' do
          expects = default_attrs.merge({ :name => 'Courier New', :size => '24' })
          is_expected.to have_attributes(expects)
        end
      end

      context 'when colored BLACK' do
        let(:cell) { { :row => 7, :col => 'A' } }

        it 'returns default font attributes (which include black)' do
          is_expected.to have_attributes(default_attrs)
        end
      end

      context 'when colored RED' do
        let(:cell) { { :row => 8, :col => 'A' } }

        it 'returns defaults plus red color' do
          expects = default_attrs.merge({ :color => '#FF0000' })
          is_expected.to have_attributes(expects)
        end
      end

      context 'when colored BLUE' do
        let(:cell) { { :row => 9, :col => 'A' } }

        it 'returns defaults plus blue color' do
          expects = default_attrs.merge({ :color => '#0066CC' })
          is_expected.to have_attributes(expects)
        end
      end

      context 'when BOLD' do
        let(:cell) { { :row => 10, :col => 'A' } }

        it 'returns defaults plus bold style' do
          # somehow in Excel, this ended up "no color" rather than black...
          expects = default_attrs.merge({ :bold? => true, :color => nil })
          is_expected.to have_attributes(expects)
        end
      end

      context 'when ITALIC' do
        let(:cell) { { :row => 11, :col => 'A' } }

        it 'returns defaults plus italic style' do
          # somehow in Excel, this ended up "no color" rather than black...
          expects = default_attrs.merge({ :italic? => true, :color => nil })
          is_expected.to have_attributes(expects)
        end
      end

      context 'when UNDERLINED' do
        let(:cell) { { :row => 12, :col => 'A' } }

        it 'returns defaults plus underlined style' do
          # somehow in Excel, this ended up "no color" rather than black...
          expects = default_attrs.merge({ :underline? => true, :color => nil })
          is_expected.to have_attributes(expects)
        end
      end

      context 'when BOLD, ITALIC, UNDERLINED, and colored PURPLE' do
        let(:cell) { { :row => 13, :col => 'A' } }

        it 'returns defaults plus bold, italic, underlined, and purple color' do
          expects = default_attrs.merge({
                                          :color      => '#666699',
                                          :bold?      => true,
                                          :italic?    => true,
                                          :underline? => true
                                        })
          is_expected.to have_attributes(expects)
        end
      end
    end

    context 'with styling set on an entire row' do
      let(:row_style) do
        default_attrs.merge({ :color => '#ED7D31', :bold? => true })
      end

      context 'when no cell styling' do
        let(:cell) { { :row => 14, :col => 'L' } }

        it 'returns the row style' do
          is_expected.to have_attributes(row_style)
        end
      end
    end

    context 'with styling set on an entire column' do
      let(:col_style) do
        default_attrs.merge({ :color => '#00FF00', :italic? => true })
      end

      context 'when no cell styling' do
        let(:cell) { { :row => 20, :col => 'K' } }

        it 'returns the column style' do
          is_expected.to have_attributes(col_style)
        end
      end
    end
  end
end
