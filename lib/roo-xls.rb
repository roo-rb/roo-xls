require 'roo'

module Roo
  CLASS_FOR_EXTENSION.merge!(
    '.xls' => 'Roo::Excel',
    '.xml' => 'Roo::Excel2003XML'
  )

  autoload :Excel, 'roo/xls/excel'
  autoload :Excel2003XML, 'roo/xls/excel_2003_xml'
end
