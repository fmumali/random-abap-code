  METHOD add_new_autofilter.
* Check for autofilter reference: new or overwrite; only one per sheet
    ro_autofilter = autofilters->add( io_sheet ) .
  ENDMETHOD.