require 'spec_helper'

describe Roo::Excel do
  let(:excel) { Roo::Excel.new('test/files/boolean.xls') }

  describe '.new' do
    it 'creates an instance' do
      expect(excel).to be_a(Roo::Excel)
    end
  end

  describe '#sheets' do
    it 'returns the sheet names of the file' do
      expect(excel.sheets).to eq(["Sheet1", "Sheet2", "Sheet3"])
    end
  end

  describe '#each_row' do
    it 'yields each row to a block' do
      expect { |b| excel.each_row(&b) }
        .to yield_successive_args(['true'], ['false'])
    end
  end

  describe '#each_row_with_index' do
    it 'yields each row to a block with the index' do
      expect { |b| excel.each_row_with_index(&b) }
        .to yield_successive_args([['true'], 1], [['false'], 2])
    end
  end
end
