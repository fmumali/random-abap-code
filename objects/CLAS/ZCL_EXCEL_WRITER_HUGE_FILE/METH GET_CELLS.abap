  METHOD get_cells.
*
* Callback method from transformation ZEXCEL_TR_SHEET
*
* The method fills the data cells for each row.
* This saves memory if there are many rows.
*
    DATA:
      lv_cell_style TYPE zexcel_cell_style.

    FIELD-SYMBOLS:
      <cell>    TYPE ty_cell,
      <content> TYPE zexcel_s_cell_data,
      <style>   TYPE zexcel_s_styles_mapping.

    CLEAR cells.

    LOOP AT worksheet->sheet_content FROM i_index ASSIGNING <content>.
      IF <content>-cell_row <> i_row.
*
*     End of row
*
        EXIT.
      ENDIF.

*
*   Determine style index
*
      IF lv_cell_style <> <content>-cell_style.
        lv_cell_style = <content>-cell_style.
        UNASSIGN <style>.
        IF lv_cell_style IS NOT INITIAL.
          READ TABLE styles_mapping ASSIGNING <style> WITH KEY guid = lv_cell_style.
        ENDIF.
      ENDIF.
*
*   Add a new cell
*
      APPEND INITIAL LINE TO cells ASSIGNING <cell>.
      <cell>-name    = <content>-cell_coords.
      <cell>-formula = <content>-cell_formula.
      <cell>-type    = <content>-data_type.
      IF <cell>-type = 's'.
        <cell>-value = me->get_shared_string_index( ip_cell_value = <content>-cell_value ).
      ELSE.
        <cell>-value = <content>-cell_value.
      ENDIF.
      IF <style> IS ASSIGNED.
        <cell>-style = <style>-style.
      ELSE.
        <cell>-style = -1.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.