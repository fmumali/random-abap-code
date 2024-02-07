  METHOD add_new_style.
* Start of deletion # issue 139 - Dateretention of cellstyles
*  CREATE OBJECT eo_style.
*  styles->add( eo_style ).
* End of deletion # issue 139 - Dateretention of cellstyles
* Start of insertion # issue 139 - Dateretention of cellstyles
* Create default style
    CREATE OBJECT eo_style
      EXPORTING
        ip_guid     = ip_guid
        io_clone_of = io_clone_of.
    styles->add( eo_style ).

    DATA: style2 TYPE zexcel_s_stylemapping.
* Copy to new representations
    style2 = stylemapping_dynamic_style( eo_style ).
    INSERT style2 INTO TABLE t_stylemapping1.
    INSERT style2 INTO TABLE t_stylemapping2.
* End of insertion # issue 139 - Dateretention of cellstyles

  ENDMETHOD.