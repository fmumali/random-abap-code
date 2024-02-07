  METHOD guid_create.

    TRY.
        rv_guid_16 = cl_system_uuid=>if_system_uuid_static~create_uuid_x16( ).
      CATCH cx_uuid_error.
    ENDTRY.

*--------------------------------------------------------------------*
* If you are on a release that does not yet have the class cl_system_uuid
* please use the following coding instead which is using the function
* call that was used before but which has been flagged as obsolete
* in newer SAP releases
*--------------------------------------------------------------------*
*
*Before ABAP 7.02:  CALL FUNCTION 'GUID_CREATE'
*Before ABAP 7.02:    IMPORTING
*Before ABAP 7.02:      ev_guid_16 = rv_guid_16.

  ENDMETHOD.