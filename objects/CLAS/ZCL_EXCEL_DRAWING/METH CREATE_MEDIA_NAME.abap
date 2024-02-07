  METHOD create_media_name.

* if media name is initial, create unique name
    CHECK media_name IS INITIAL.

    index = ip_index.
    CONCATENATE me->type index INTO media_name.
    CONDENSE media_name NO-GAPS.
  ENDMETHOD.