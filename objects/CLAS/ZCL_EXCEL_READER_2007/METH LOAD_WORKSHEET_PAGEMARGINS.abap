  METHOD load_worksheet_pagemargins.

    TYPES: BEGIN OF lty_page_margins,
             footer TYPE string,
             header TYPE string,
             bottom TYPE string,
             top    TYPE string,
             right  TYPE string,
             left   TYPE string,
           END OF lty_page_margins.

    DATA:lo_ixml_pagemargins_elem TYPE REF TO if_ixml_element,
         ls_pagemargins           TYPE lty_page_margins.


    lo_ixml_pagemargins_elem = io_ixml_worksheet->find_from_name_ns( name = 'pageMargins' uri = namespace-main ).
    IF lo_ixml_pagemargins_elem IS NOT INITIAL.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_pagemargins_elem
                                   CHANGING
                                     cp_structure = ls_pagemargins ).
      io_worksheet->sheet_setup->margin_bottom = zcl_excel_common=>excel_string_to_number( ls_pagemargins-bottom ).
      io_worksheet->sheet_setup->margin_footer = zcl_excel_common=>excel_string_to_number( ls_pagemargins-footer ).
      io_worksheet->sheet_setup->margin_header = zcl_excel_common=>excel_string_to_number( ls_pagemargins-header ).
      io_worksheet->sheet_setup->margin_left   = zcl_excel_common=>excel_string_to_number( ls_pagemargins-left   ).
      io_worksheet->sheet_setup->margin_right  = zcl_excel_common=>excel_string_to_number( ls_pagemargins-right  ).
      io_worksheet->sheet_setup->margin_top    = zcl_excel_common=>excel_string_to_number( ls_pagemargins-top    ).
    ENDIF.

  ENDMETHOD.