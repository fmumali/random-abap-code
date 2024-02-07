  METHOD register_new_style.


    me->add( io_style ).
    ep_style_code = me->size( ) - 1. "style count starts from 0
  ENDMETHOD.