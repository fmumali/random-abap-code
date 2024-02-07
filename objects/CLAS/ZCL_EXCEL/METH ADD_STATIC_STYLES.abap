  METHOD add_static_styles.
    " # issue 139
    FIELD-SYMBOLS: <style1> LIKE LINE OF t_stylemapping1,
                   <style2> LIKE LINE OF t_stylemapping2.
    DATA: style TYPE REF TO zcl_excel_style.

    LOOP AT me->t_stylemapping1 ASSIGNING <style1> WHERE added_to_iterator IS INITIAL.
      READ TABLE me->t_stylemapping2 ASSIGNING <style2> WITH TABLE KEY guid = <style1>-guid.
      CHECK sy-subrc = 0.  " Should always be true since these tables are being filled parallel

      style = me->add_new_style( <style1>-guid ).

      zcl_excel_common=>recursive_struct_to_class( EXPORTING i_source  = <style1>-complete_style
                                                             i_sourcex = <style1>-complete_stylex
                                                   CHANGING  e_target  = style ).

    ENDLOOP.
  ENDMETHOD.