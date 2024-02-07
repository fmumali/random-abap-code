  METHOD stylemapping_dynamic_style.
    " # issue 139
    eo_style2-dynamic_style_guid  = ip_style->get_guid( ).
    eo_style2-guid                = eo_style2-dynamic_style_guid.
    eo_style2-added_to_iterator   = abap_true.

* don't care about attributes here, since this data may change
* dynamically

  ENDMETHOD.