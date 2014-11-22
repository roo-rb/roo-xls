require 'roo'

module Roo
  autoload :Excel, 'roo/xls/excel'
  autoload :Excel2003XML, 'roo/xls/excel_2003_xml'

  CLASS_FOR_EXTENSION.merge!(
    xls: Roo::Excel,
    xml: Roo::Excel2003XML
  )
end
