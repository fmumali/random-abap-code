  METHOD add_new_comment.
    CREATE OBJECT eo_comment.

    comments->add( eo_comment ).
  ENDMETHOD.