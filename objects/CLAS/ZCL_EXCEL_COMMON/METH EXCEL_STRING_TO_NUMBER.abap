  METHOD excel_string_to_number.

* If we encounter anything more complicated in EXCEL we might have to extend this
* But currently this works fine - even for numbers in scientific notation

    ep_value = ip_value.

  ENDMETHOD.