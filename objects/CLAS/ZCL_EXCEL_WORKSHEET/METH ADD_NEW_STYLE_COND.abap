  METHOD add_new_style_cond.
    CREATE OBJECT eo_style_cond EXPORTING ip_dimension_range = ip_dimension_range.
    styles_cond->add( eo_style_cond ).
  ENDMETHOD.                    "ADD_NEW_STYLE_COND