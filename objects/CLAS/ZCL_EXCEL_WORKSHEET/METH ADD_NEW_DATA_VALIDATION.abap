  METHOD add_new_data_validation.

    CREATE OBJECT eo_data_validation.
    data_validations->add( eo_data_validation ).
  ENDMETHOD.                    "ADD_NEW_DATA_VALIDATION