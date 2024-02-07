  METHOD create_ax.
    DATA ls_ax TYPE s_ax.
    ls_ax-type = ip_type.

    IF ip_type = c_catax.
      IF ip_axid IS SUPPLIED.
        ls_ax-axid = ip_axid.
      ELSE.
        ls_ax-axid = '1'.
      ENDIF.
      IF ip_orientation IS SUPPLIED.
        ls_ax-orientation = ip_orientation.
      ELSE.
        ls_ax-orientation = 'minMax'.
      ENDIF.
      IF ip_delete IS SUPPLIED.
        ls_ax-delete = ip_delete.
      ELSE.
        ls_ax-delete = '0'.
      ENDIF.
      IF ip_axpos IS SUPPLIED.
        ls_ax-axpos = ip_axpos.
      ELSE.
        ls_ax-axpos = 'b'.
      ENDIF.
      IF ip_formatcode IS SUPPLIED.
        ls_ax-formatcode = ip_formatcode.
      ELSE.
        ls_ax-formatcode = 'General'.
      ENDIF.
      IF ip_sourcelinked IS SUPPLIED.
        ls_ax-sourcelinked = ip_sourcelinked.
      ELSE.
        ls_ax-sourcelinked = '1'.
      ENDIF.
      IF ip_majortickmark IS SUPPLIED.
        ls_ax-majortickmark = ip_majortickmark.
      ELSE.
        ls_ax-majortickmark = 'out'.
      ENDIF.
      IF ip_minortickmark IS SUPPLIED.
        ls_ax-minortickmark = ip_minortickmark.
      ELSE.
        ls_ax-minortickmark = 'none'.
      ENDIF.
      IF ip_ticklblpos IS SUPPLIED.
        ls_ax-ticklblpos = ip_ticklblpos.
      ELSE.
        ls_ax-ticklblpos = 'nextTo'.
      ENDIF.
      IF ip_crossax IS SUPPLIED.
        ls_ax-crossax = ip_crossax.
      ELSE.
        ls_ax-crossax = '2'.
      ENDIF.
      IF ip_crosses IS SUPPLIED.
        ls_ax-crosses = ip_crosses.
      ELSE.
        ls_ax-crosses = 'autoZero'.
      ENDIF.
      IF ip_auto IS SUPPLIED.
        ls_ax-auto = ip_auto.
      ELSE.
        ls_ax-auto = '1'.
      ENDIF.
      IF ip_lblalgn IS SUPPLIED.
        ls_ax-lblalgn = ip_lblalgn.
      ELSE.
        ls_ax-lblalgn = 'ctr'.
      ENDIF.
      IF ip_lbloffset IS SUPPLIED.
        ls_ax-lbloffset = ip_lbloffset.
      ELSE.
        ls_ax-lbloffset = '100'.
      ENDIF.
      IF ip_nomultilvllbl IS SUPPLIED.
        ls_ax-nomultilvllbl = ip_nomultilvllbl.
      ELSE.
        ls_ax-nomultilvllbl = '0'.
      ENDIF.
    ELSEIF ip_type = c_valax.
      IF ip_axid IS SUPPLIED.
        ls_ax-axid = ip_axid.
      ELSE.
        ls_ax-axid = '2'.
      ENDIF.
      IF ip_orientation IS SUPPLIED.
        ls_ax-orientation = ip_orientation.
      ELSE.
        ls_ax-orientation = 'minMax'.
      ENDIF.
      IF ip_delete IS SUPPLIED.
        ls_ax-delete = ip_delete.
      ELSE.
        ls_ax-delete = '0'.
      ENDIF.
      IF ip_axpos IS SUPPLIED.
        ls_ax-axpos = ip_axpos.
      ELSE.
        ls_ax-axpos = 'l'.
      ENDIF.
      IF ip_formatcode IS SUPPLIED.
        ls_ax-formatcode = ip_formatcode.
      ELSE.
        ls_ax-formatcode = 'General'.
      ENDIF.
      IF ip_sourcelinked IS SUPPLIED.
        ls_ax-sourcelinked = ip_sourcelinked.
      ELSE.
        ls_ax-sourcelinked = '1'.
      ENDIF.
      IF ip_majortickmark IS SUPPLIED.
        ls_ax-majortickmark = ip_majortickmark.
      ELSE.
        ls_ax-majortickmark = 'out'.
      ENDIF.
      IF ip_minortickmark IS SUPPLIED.
        ls_ax-minortickmark = ip_minortickmark.
      ELSE.
        ls_ax-minortickmark = 'none'.
      ENDIF.
      IF ip_ticklblpos IS SUPPLIED.
        ls_ax-ticklblpos = ip_ticklblpos.
      ELSE.
        ls_ax-ticklblpos = 'nextTo'.
      ENDIF.
      IF ip_crossax IS SUPPLIED.
        ls_ax-crossax = ip_crossax.
      ELSE.
        ls_ax-crossax = '1'.
      ENDIF.
      IF ip_crosses IS SUPPLIED.
        ls_ax-crosses = ip_crosses.
      ELSE.
        ls_ax-crosses = 'autoZero'.
      ENDIF.
      IF ip_crossbetween IS SUPPLIED.
        ls_ax-crossbetween = ip_crossbetween.
      ELSE.
        ls_ax-crossbetween = 'between'.
      ENDIF.
    ENDIF.

    APPEND ls_ax TO me->axes.
    SORT me->axes BY axid ASCENDING.
  ENDMETHOD.