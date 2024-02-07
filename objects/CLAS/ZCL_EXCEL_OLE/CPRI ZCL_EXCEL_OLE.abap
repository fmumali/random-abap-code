  PRIVATE SECTION.

    CLASS-METHODS close_document.
    CLASS-METHODS error_doi.

    CLASS-DATA: lo_spreadsheet TYPE REF TO i_oi_spreadsheet,
                lo_control     TYPE REF TO i_oi_container_control,
                lo_proxy       TYPE REF TO i_oi_document_proxy,
                lo_error       TYPE REF TO i_oi_error,
                lc_retcode     TYPE        soi_ret_string.
