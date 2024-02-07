  METHOD bind_alv_ole2.

    CALL METHOD ('ZCL_EXCEL_OLE')=>('BIND_ALV_OLE2')
      EXPORTING
        i_document_url          = i_document_url
        i_xls                   = i_xls
        i_save_path             = i_save_path
        io_alv                  = io_alv
        it_listheader           = it_listheader
        i_top                   = i_top
        i_left                  = i_left
        i_columns_header        = i_columns_header
        i_columns_autofit       = i_columns_autofit
        i_format_col_header     = i_format_col_header
        i_format_subtotal       = i_format_subtotal
        i_format_total          = i_format_total
      EXCEPTIONS
        miss_guide              = 1
        ex_transfer_kkblo_error = 2
        fatal_error             = 3
        inv_data_range          = 4
        dim_mismatch_vkey       = 5
        dim_mismatch_sema       = 6
        error_in_sema           = 7
        OTHERS                  = 8.
    IF sy-subrc <> 0.
      CASE sy-subrc.
        WHEN 1. RAISE miss_guide.
        WHEN 2. RAISE ex_transfer_kkblo_error.
        WHEN 3. RAISE fatal_error.
        WHEN 4. RAISE inv_data_range.
        WHEN 5. RAISE dim_mismatch_vkey.
        WHEN 6. RAISE dim_mismatch_sema.
        WHEN 7. RAISE error_in_sema.
      ENDCASE.
    ENDIF.

  ENDMETHOD.                    "BIND_ALV_OLE2