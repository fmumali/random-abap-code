  METHOD load_style_num_formats.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Explain gaps in predefined formats
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-25
*              - ...
* changes: renaming variables and types to naming conventions
*          adding comments to explain what we are trying to achieve
*          aligning code
*--------------------------------------------------------------------*
    DATA: lo_node_numfmt TYPE REF TO if_ixml_element,
          ls_num_format  TYPE zcl_excel_style_number_format=>t_num_format.

*--------------------------------------------------------------------*
* We need a table of used numberformats to build up our styles
* there are two kinds of numberformats
* §1 built-in numberformats
* §2 and those that have been explicitly added by the createor of the excel-file
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* §1 built-in numberformats
*--------------------------------------------------------------------*
    ep_num_formats = zcl_excel_style_number_format=>mt_built_in_num_formats.

*--------------------------------------------------------------------*
* §2   Get non-internal numberformats that are found in the file explicitly

*         Following is an example how this part of a file could be set up
*         <numFmts count="1">
*             <numFmt formatCode="#,###,###,###,##0.00" numFmtId="164"/>
*         </numFmts>
*--------------------------------------------------------------------*
    lo_node_numfmt ?= ip_xml->find_from_name_ns( name = 'numFmt' uri = namespace-main ).
    WHILE lo_node_numfmt IS BOUND.

      CLEAR ls_num_format.

      CREATE OBJECT ls_num_format-format.
      ls_num_format-format->format_code = lo_node_numfmt->get_attribute( 'formatCode' ).
      ls_num_format-id                  = lo_node_numfmt->get_attribute( 'numFmtId' ).
      INSERT ls_num_format INTO TABLE ep_num_formats.

      lo_node_numfmt                          ?= lo_node_numfmt->get_next( ).

    ENDWHILE.


  ENDMETHOD.