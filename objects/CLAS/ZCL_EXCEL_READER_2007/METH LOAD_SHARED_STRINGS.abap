  METHOD load_shared_strings.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Support partial formatting of strings in cells
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-11
*              - ...
* changes: renaming variables to naming conventions
*          renaming variables to indicate what they are used for
*          aligning code
*          adding comments to explain what we are trying to achieve
*          rewriting code for better readibility
*--------------------------------------------------------------------*



    DATA:
      lo_shared_strings_xml TYPE REF TO if_ixml_document,
      lo_node_si            TYPE REF TO if_ixml_element,
      lo_node_si_child      TYPE REF TO if_ixml_element,
      lo_node_r_child_t     TYPE REF TO if_ixml_element,
      lo_node_r_child_rpr   TYPE REF TO if_ixml_element,
      lo_font               TYPE REF TO zcl_excel_style_font,
      ls_rtf                TYPE zexcel_s_rtf,
      lv_current_offset     TYPE int2,
      lv_tag_name           TYPE string,
      lv_node_value         TYPE string.

    FIELD-SYMBOLS: <ls_shared_string>           LIKE LINE OF me->shared_strings.

*--------------------------------------------------------------------*

* §1  Parse shared strings file and get into internal table
*   So far I have encountered 2 ways how a string can be represented in the shared strings file
*   §1.1 - "simple" strings
*   §1.2 - rich text formatted strings

*     Following is an example how this file could be set up; 2 strings in simple formatting, 3rd string rich textformatted


*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <sst uniqueCount="6" count="6" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
*            <si>
*                <t>This is a teststring 1</t>
*            </si>
*            <si>
*                <t>This is a teststring 2</t>
*            </si>
*            <si>
*                <r>
*                  <t>T</t>
*                </r>
*                <r>
*                    <rPr>
*                        <sz val="11"/>
*                        <color rgb="FFFF0000"/>
*                        <rFont val="Calibri"/>
*                        <family val="2"/>
*                        <scheme val="minor"/>
*                    </rPr>
*                    <t xml:space="preserve">his is a </t>
*                </r>
*                <r>
*                    <rPr>
*                        <sz val="11"/>
*                        <color theme="1"/>
*                        <rFont val="Calibri"/>
*                        <family val="2"/>
*                        <scheme val="minor"/>
*                    </rPr>
*                    <t>teststring 3</t>
*                </r>
*            </si>
*        </sst>
*--------------------------------------------------------------------*

    lo_shared_strings_xml = me->get_ixml_from_zip_archive( i_filename     = ip_path
                                                           is_normalizing = space ).  " NO!!! normalizing - otherwise leading blanks will be omitted and that is not really desired for the stringtable
    lo_node_si ?= lo_shared_strings_xml->find_from_name_ns( name = 'si' uri = namespace-main ).
    WHILE lo_node_si IS BOUND.

      APPEND INITIAL LINE TO me->shared_strings ASSIGNING <ls_shared_string>.            " Each <si>-entry in the xml-file must lead to an entry in our stringtable
      lo_node_si_child ?= lo_node_si->get_first_child( ).
      IF lo_node_si_child IS BOUND.
        lv_tag_name = lo_node_si_child->get_name( ).
        IF lv_tag_name = 't'.
*--------------------------------------------------------------------*
*   §1.1 - "simple" strings
*                Example:  see above
*--------------------------------------------------------------------*
          <ls_shared_string>-value = unescape_string_value( lo_node_si_child->get_value( ) ).
        ELSE.
*--------------------------------------------------------------------*
*   §1.2 - rich text formatted strings
*       it is sufficient to strip the <t>...</t> tag from each <r>-tag and concatenate these
*       as long as rich text formatting is not supported (2do§1) ignore all info about formatting
*                Example:  see above
*--------------------------------------------------------------------*
          CLEAR: lv_current_offset.
          WHILE lo_node_si_child IS BOUND.                                             " actually these children of <si> are <r>-tags
            CLEAR: ls_rtf.

            " extracting rich text formating data
            lo_node_r_child_rpr ?= lo_node_si_child->find_from_name_ns( name = 'rPr' uri = namespace-main ).
            IF lo_node_r_child_rpr IS BOUND.
              lo_font = load_style_font( lo_node_r_child_rpr ).
              ls_rtf-font = lo_font->get_structure( ).
            ENDIF.
            ls_rtf-offset = lv_current_offset.
            " extract the <t>...</t> part of each <r>-tag
            lo_node_r_child_t ?= lo_node_si_child->find_from_name_ns( name = 't' uri = namespace-main ).
            IF lo_node_r_child_t IS BOUND.
              lv_node_value = unescape_string_value( lo_node_r_child_t->get_value( ) ).
              CONCATENATE <ls_shared_string>-value lv_node_value INTO <ls_shared_string>-value RESPECTING BLANKS.
              ls_rtf-length = strlen( lv_node_value ).
            ENDIF.

            lv_current_offset = strlen( <ls_shared_string>-value ).
            APPEND ls_rtf TO <ls_shared_string>-rtf.

            lo_node_si_child ?= lo_node_si_child->get_next( ).

          ENDWHILE.
        ENDIF.
      ENDIF.

      lo_node_si ?= lo_node_si->get_next( ).
    ENDWHILE.

  ENDMETHOD.