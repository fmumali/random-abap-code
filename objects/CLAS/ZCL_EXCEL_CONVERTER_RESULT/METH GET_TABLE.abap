  METHOD get_table.
    DATA: lo_object   TYPE REF TO object,
          ls_seoclass TYPE seoclass,
          l_method    TYPE string.

    SELECT SINGLE * INTO ls_seoclass
      FROM seoclass
      WHERE clsname = 'IF_SALV_BS_DATA_SOURCE'.

    IF sy-subrc = 0.
      l_method = 'GET_TABLE_REF'.
      lo_object ?= io_object.
      CALL METHOD lo_object->(l_method)
        RECEIVING
          value = ro_data.
    ELSE.
      l_method = 'GET_REF_TO_TABLE'.
      lo_object ?= io_object.
      CALL METHOD lo_object->(l_method)
        RECEIVING
          value = ro_data.
    ENDIF.

  ENDMETHOD.