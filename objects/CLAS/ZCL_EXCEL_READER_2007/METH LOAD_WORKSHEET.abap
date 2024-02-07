  METHOD load_worksheet.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Header/footer
*
*                Please don't just delete these ToDos if they are not
*                needed but leave a comment that states this
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,
*              - ...
* changes: renaming variables to naming conventions
*          aligning code                                            (started)
*          add a list of open ToDos here
*          adding comments to explain what we are trying to achieve (started)
*--------------------------------------------------------------------*
* issue #345 - Dump on small pagemargins
*              Took the chance to modularize this very long method
*              by extracting the code that needed correction into
*              own method ( load_worksheet_pagemargins )
*--------------------------------------------------------------------*
    TYPES: BEGIN OF lty_cell,
             r TYPE string,
             t TYPE string,
             s TYPE string,
           END OF lty_cell.

    TYPES: BEGIN OF lty_column,
             min          TYPE string,
             max          TYPE string,
             width        TYPE f,
             customwidth  TYPE string,
             style        TYPE string,
             bestfit      TYPE string,
             collapsed    TYPE string,
             hidden       TYPE string,
             outlinelevel TYPE string,
           END OF lty_column.

    TYPES: BEGIN OF lty_sheetview,
             showgridlines            TYPE zexcel_show_gridlines,
             tabselected              TYPE string,
             zoomscale                TYPE string,
             zoomscalenormal          TYPE string,
             zoomscalepagelayoutview  TYPE string,
             zoomscalesheetlayoutview TYPE string,
             workbookviewid           TYPE string,
             showrowcolheaders        TYPE string,
             righttoleft              TYPE string,
             topleftcell       TYPE string,
           END OF lty_sheetview.

    TYPES: BEGIN OF lty_mergecell,
             ref TYPE string,
           END OF lty_mergecell.

    TYPES: BEGIN OF lty_row,
             r            TYPE string,
             customheight TYPE string,
             ht           TYPE f,
             spans        TYPE string,
             thickbot     TYPE string,
             customformat TYPE string,
             thicktop     TYPE string,
             collapsed    TYPE string,
             hidden       TYPE string,
             outlinelevel TYPE string,
           END OF lty_row.

    TYPES: BEGIN OF lty_page_setup,
             id          TYPE string,
             orientation TYPE string,
             scale       TYPE string,
             fittoheight TYPE string,
             fittowidth  TYPE string,
             papersize   TYPE string,
             paperwidth  TYPE string,
             paperheight TYPE string,
           END OF lty_page_setup.

    TYPES: BEGIN OF lty_sheetformatpr,
             customheight     TYPE string,
             defaultrowheight TYPE string,
             customwidth      TYPE string,
             defaultcolwidth  TYPE string,
           END OF lty_sheetformatpr.

    TYPES: BEGIN OF lty_headerfooter,
             alignwithmargins TYPE string,
             differentoddeven TYPE string,
           END OF lty_headerfooter.

    TYPES: BEGIN OF lty_tabcolor,
             rgb   TYPE string,
             theme TYPE string,
           END OF lty_tabcolor.

    TYPES: BEGIN OF lty_datavalidation,
             type             TYPE zexcel_data_val_type,
             allowblank       TYPE flag,
             showinputmessage TYPE flag,
             showerrormessage TYPE flag,
             showdropdown     TYPE flag,
             operator         TYPE zexcel_data_val_operator,
             formula1         TYPE zexcel_validation_formula1,
             formula2         TYPE zexcel_validation_formula1,
             sqref            TYPE string,
             cell_column      TYPE zexcel_cell_column_alpha,
             cell_column_to   TYPE zexcel_cell_column_alpha,
             cell_row         TYPE zexcel_cell_row,
             cell_row_to      TYPE zexcel_cell_row,
             error            TYPE string,
             errortitle       TYPE string,
             prompt           TYPE string,
             prompttitle      TYPE string,
             errorstyle       TYPE zexcel_data_val_error_style,
           END OF lty_datavalidation.



    CONSTANTS: lc_xml_attr_true     TYPE string VALUE 'true',
               lc_xml_attr_true_int TYPE string VALUE '1',
               lc_rel_drawing       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
               lc_rel_hyperlink     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
               lc_rel_comments      TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
               lc_rel_printer       TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'.
    CONSTANTS lc_rel_table TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'.

    DATA: lo_ixml_worksheet           TYPE REF TO if_ixml_document,
          lo_ixml_cells               TYPE REF TO if_ixml_node_collection,
          lo_ixml_iterator            TYPE REF TO if_ixml_node_iterator,
          lo_ixml_iterator2           TYPE REF TO if_ixml_node_iterator,
          lo_ixml_row_elem            TYPE REF TO if_ixml_element,
          lo_ixml_cell_elem           TYPE REF TO if_ixml_element,
          ls_cell                     TYPE lty_cell,
          lv_index                    TYPE i,
          lv_index_temp               TYPE i,
          lo_ixml_value_elem          TYPE REF TO if_ixml_element,
          lo_ixml_formula_elem        TYPE REF TO if_ixml_element,
          lv_cell_value               TYPE zexcel_cell_value,
          lv_cell_formula             TYPE zexcel_cell_formula,
          lv_cell_column              TYPE zexcel_cell_column_alpha,
          lv_cell_row                 TYPE zexcel_cell_row,
          lo_excel_style              TYPE REF TO zcl_excel_style,
          lv_style_guid               TYPE zexcel_cell_style,

          lo_ixml_imension_elem       TYPE REF TO if_ixml_element, "#+234
          lv_dimension_range          TYPE string,  "#+234

          lo_ixml_sheetview_elem      TYPE REF TO if_ixml_element,
          ls_sheetview                TYPE lty_sheetview,
          lo_ixml_pane_elem           TYPE REF TO if_ixml_element,
          ls_excel_pane               TYPE zexcel_pane,
          lv_pane_cell_row            TYPE zexcel_cell_row,
          lv_pane_cell_col_a          TYPE zexcel_cell_column_alpha,
          lv_pane_cell_col            TYPE zexcel_cell_column,

          lo_ixml_mergecells          TYPE REF TO if_ixml_node_collection,
          lo_ixml_mergecell_elem      TYPE REF TO if_ixml_element,
          ls_mergecell                TYPE lty_mergecell,
          lv_merge_column_start       TYPE zexcel_cell_column_alpha,
          lv_merge_column_end         TYPE zexcel_cell_column_alpha,
          lv_merge_row_start          TYPE zexcel_cell_row,
          lv_merge_row_end            TYPE zexcel_cell_row,

          lo_ixml_sheetformatpr_elem  TYPE REF TO if_ixml_element,
          ls_sheetformatpr            TYPE lty_sheetformatpr,
          lv_height                   TYPE f,

          lo_ixml_headerfooter_elem   TYPE REF TO if_ixml_element,
          ls_headerfooter             TYPE lty_headerfooter,
          ls_odd_header               TYPE zexcel_s_worksheet_head_foot,
          ls_odd_footer               TYPE zexcel_s_worksheet_head_foot,
          ls_even_header              TYPE zexcel_s_worksheet_head_foot,
          ls_even_footer              TYPE zexcel_s_worksheet_head_foot,
          lo_ixml_hf_value_elem       TYPE REF TO if_ixml_element,

          lo_ixml_pagesetup_elem      TYPE REF TO if_ixml_element,
          lo_ixml_sheetpr             TYPE REF TO if_ixml_element,
          lv_fit_to_page              TYPE string,
          ls_pagesetup                TYPE lty_page_setup,

          lo_ixml_columns             TYPE REF TO if_ixml_node_collection,
          lo_ixml_column_elem         TYPE REF TO if_ixml_element,
          ls_column                   TYPE lty_column,
          lv_column_alpha             TYPE zexcel_cell_column_alpha,
          lo_column                   TYPE REF TO zcl_excel_column,
          lv_outline_level            TYPE int4,

          lo_ixml_tabcolor            TYPE REF TO if_ixml_element,
          ls_tabcolor                 TYPE lty_tabcolor,
          ls_excel_s_tabcolor         TYPE zexcel_s_tabcolor,

          lo_ixml_rows                TYPE REF TO if_ixml_node_collection,
          ls_row                      TYPE lty_row,
          lv_max_col                  TYPE i,     "for use with SPANS element
*              lv_min_col                     TYPE i,     "for use with SPANS element                    " not in use currently
          lv_max_col_s                TYPE char10,     "for use with SPANS element
          lv_min_col_s                TYPE char10,     "for use with SPANS element
          lo_row                      TYPE REF TO zcl_excel_row,
*---    End of current code aligning -------------------------------------------------------------------

          lv_path                     TYPE string,
          lo_ixml_node                TYPE REF TO if_ixml_element,
          ls_relationship             TYPE t_relationship,
          lo_ixml_rels_worksheet      TYPE REF TO if_ixml_document,
          lv_rels_worksheet_path      TYPE string,
          lv_stripped_name            TYPE chkfile,
          lv_dirname                  TYPE string,

          lt_external_hyperlinks      TYPE gtt_external_hyperlinks,
          ls_external_hyperlink       LIKE LINE OF lt_external_hyperlinks,

          lo_ixml_datavalidations     TYPE REF TO if_ixml_node_collection,
          lo_ixml_datavalidation_elem TYPE REF TO if_ixml_element,
          ls_datavalidation           TYPE lty_datavalidation,
          lo_data_validation          TYPE REF TO zcl_excel_data_validation,
          lv_datavalidation_range     TYPE string,
          lt_datavalidation_range     TYPE TABLE OF string,
          lt_rtf                      TYPE zexcel_t_rtf,
          ex                          TYPE REF TO cx_root.
    DATA lt_tables TYPE t_tables.
    DATA ls_table TYPE t_table.

    FIELD-SYMBOLS:
      <ls_shared_string> TYPE t_shared_string.

*--------------------------------------------------------------------*
* §2  We need to read the the file "\\_rels\.rels" because it tells
*     us where in this folder structure the data for the workbook
*     is located in the xlsx zip-archive
*
*     The xlsx Zip-archive has generally the following folder structure:
*       <root> |
*              |-->  _rels
*              |-->  doc_Props
*              |-->  xl |
*                       |-->  _rels
*                       |-->  theme
*                       |-->  worksheets
*--------------------------------------------------------------------*

    " Read Workbook Relationships
    CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = ip_path
      IMPORTING
        stripped_name = lv_stripped_name
        file_path     = lv_dirname.
    CONCATENATE lv_dirname '_rels/' lv_stripped_name '.rels'
      INTO lv_rels_worksheet_path.
    TRY.                                                                          " +#222  _rels/xxx.rels might not be present.  If not found there can be no drawings --> just ignore this section
        lo_ixml_rels_worksheet = me->get_ixml_from_zip_archive( lv_rels_worksheet_path ).
        lo_ixml_node ?= lo_ixml_rels_worksheet->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ).
      CATCH zcx_excel.                            "#EC NO_HANDLER +#222
        " +#222   No errorhandling necessary - node will be unbound if error occurs
    ENDTRY.                                                   " +#222
    WHILE lo_ixml_node IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_node
                                   CHANGING
                                     cp_structure = ls_relationship ).
      CONCATENATE lv_dirname ls_relationship-target INTO lv_path.
      lv_path = resolve_path( lv_path ).

      CASE ls_relationship-type.
        WHEN lc_rel_drawing.
          " Read Drawings
* Issue # 339       Not all drawings are in the path mentioned below.
*                   Some Excel elements like textfields (which we don't support ) have a drawing-part in the relationsships
*                   but no "xl/drawings/_rels/drawing____.xml.rels" part.
*                   Since we don't support these there is no need to read them.  Catching exceptions thrown
*                   in the "load_worksheet_drawing" shouldn't lead to an abortion of the reading
          TRY.
              me->load_worksheet_drawing( ip_path      = lv_path
                                        io_worksheet = io_worksheet ).
            CATCH zcx_excel. "--> then ignore it
          ENDTRY.

        WHEN lc_rel_printer.
          " Read Printer settings

        WHEN lc_rel_hyperlink.
          MOVE-CORRESPONDING ls_relationship TO ls_external_hyperlink.
          INSERT ls_external_hyperlink INTO TABLE lt_external_hyperlinks.

        WHEN lc_rel_comments.
          TRY.
              me->load_comments( ip_path      = lv_path
                                 io_worksheet = io_worksheet ).
            CATCH zcx_excel.
          ENDTRY.

        WHEN lc_rel_table.
          MOVE-CORRESPONDING ls_relationship TO ls_table.
          INSERT ls_table INTO TABLE lt_tables.

        WHEN OTHERS.
      ENDCASE.

      lo_ixml_node ?= lo_ixml_node->get_next( ).
    ENDWHILE.


    lo_ixml_worksheet = me->get_ixml_from_zip_archive( ip_path ).


    lo_ixml_tabcolor ?= lo_ixml_worksheet->find_from_name_ns( name = 'tabColor' uri = namespace-main ).
    IF lo_ixml_tabcolor IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_tabcolor
                                  CHANGING
                                    cp_structure = ls_tabcolor ).
* Theme not supported yet
      IF ls_tabcolor-rgb IS NOT INITIAL.
        ls_excel_s_tabcolor-rgb = ls_tabcolor-rgb.
        io_worksheet->set_tabcolor( ls_excel_s_tabcolor ).
      ENDIF.
    ENDIF.

    " Read tables (must be done before loading sheet contents)
    TRY.
        me->load_worksheet_tables( io_ixml_worksheet = lo_ixml_worksheet
                                   io_worksheet      = io_worksheet
                                   iv_dirname        = lv_dirname
                                   it_tables         = lt_tables ).
      CATCH zcx_excel. " Ignore reading errors - pass everything we were able to identify
    ENDTRY.

    " Sheet contents
    lo_ixml_rows = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'row' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_rows->create_iterator( ).
    lo_ixml_row_elem ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml_row_elem IS BOUND.

      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_row_elem
                                   CHANGING
                                     cp_structure = ls_row ).
      SPLIT ls_row-spans AT ':' INTO lv_min_col_s lv_max_col_s.
      lv_index = lv_max_col_s.
      IF lv_index > lv_max_col.
        lv_max_col = lv_index.
      ENDIF.
      lv_cell_row = ls_row-r.
      lo_row = io_worksheet->get_row( lv_cell_row ).
      IF ls_row-customheight = '1'.
        lo_row->set_row_height( ip_row_height = ls_row-ht ip_custom_height = abap_true ).
      ELSEIF ls_row-ht > 0.
        lo_row->set_row_height( ip_row_height = ls_row-ht ip_custom_height = abap_false ).
      ENDIF.

      IF   ls_row-collapsed = lc_xml_attr_true
        OR ls_row-collapsed = lc_xml_attr_true_int.
        lo_row->set_collapsed( abap_true ).
      ENDIF.

      IF   ls_row-hidden = lc_xml_attr_true
        OR ls_row-hidden = lc_xml_attr_true_int.
        lo_row->set_visible( abap_false ).
      ENDIF.

      IF ls_row-outlinelevel > ''.
*        outline_level = condense( row-outlineLevel ).  "For basis 7.02 and higher
        CONDENSE  ls_row-outlinelevel.
        lv_outline_level = ls_row-outlinelevel.
        IF lv_outline_level > 0.
          lo_row->set_outline_level( lv_outline_level ).
        ENDIF.
      ENDIF.

      lo_ixml_cells = lo_ixml_row_elem->get_elements_by_tag_name_ns( name = 'c' uri = namespace-main ).
      lo_ixml_iterator2 = lo_ixml_cells->create_iterator( ).
      lo_ixml_cell_elem ?= lo_ixml_iterator2->get_next( ).
      WHILE lo_ixml_cell_elem IS BOUND.
        CLEAR: lv_cell_value,
               lv_cell_formula,
               lv_style_guid.

        fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_cell_elem CHANGING cp_structure = ls_cell ).

        lo_ixml_value_elem = lo_ixml_cell_elem->find_from_name_ns( name = 'v' uri = namespace-main ).

        CASE ls_cell-t.
          WHEN 's'. " String values are stored as index in shared string table
            IF lo_ixml_value_elem IS BOUND.
              lv_index = lo_ixml_value_elem->get_value( ) + 1.
              READ TABLE shared_strings ASSIGNING <ls_shared_string> INDEX lv_index.
              IF sy-subrc = 0.
                lv_cell_value = <ls_shared_string>-value.
                lt_rtf = <ls_shared_string>-rtf.
              ENDIF.
            ENDIF.
          WHEN 'inlineStr'. " inlineStr values are kept in special node
            lo_ixml_value_elem = lo_ixml_cell_elem->find_from_name_ns( name = 'is' uri = namespace-main ).
            IF lo_ixml_value_elem IS BOUND.
              lv_cell_value = lo_ixml_value_elem->get_value( ).
            ENDIF.
          WHEN OTHERS. "other types are stored directly
            IF lo_ixml_value_elem IS BOUND.
              lv_cell_value = lo_ixml_value_elem->get_value( ).
            ENDIF.
        ENDCASE.

        CLEAR lv_style_guid.
        "read style based on index
        IF ls_cell-s IS NOT INITIAL.
          lv_index = ls_cell-s + 1.
          READ TABLE styles INTO lo_excel_style INDEX lv_index.
          IF sy-subrc = 0.
            lv_style_guid = lo_excel_style->get_guid( ).
          ENDIF.
        ENDIF.

        lo_ixml_formula_elem = lo_ixml_cell_elem->find_from_name_ns( name = 'f' uri = namespace-main ).
        IF lo_ixml_formula_elem IS BOUND.
          lv_cell_formula = lo_ixml_formula_elem->get_value( ).
*--------------------------------------------------------------------*
* Begin of insertion issue#284 - Copied formulae not
*--------------------------------------------------------------------*
          DATA: BEGIN OF ls_formula_attributes,
                  ref TYPE string,
                  si  TYPE i,
                  t   TYPE string,
                END OF ls_formula_attributes,
                ls_ref_formula TYPE ty_ref_formulae.

          fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_formula_elem CHANGING cp_structure = ls_formula_attributes ).
          IF ls_formula_attributes-t = 'shared'.
            zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING
                                                                i_columnrow = ls_cell-r
                                                              IMPORTING
                                                                e_column    = lv_cell_column
                                                                e_row       = lv_cell_row ).

            TRY.
                CLEAR ls_ref_formula.
                ls_ref_formula-sheet     = io_worksheet.
                ls_ref_formula-row       = lv_cell_row.
                ls_ref_formula-column    = zcl_excel_common=>convert_column2int( lv_cell_column ).
                ls_ref_formula-si        = ls_formula_attributes-si.
                ls_ref_formula-ref       = ls_formula_attributes-ref.
                ls_ref_formula-formula   = lv_cell_formula.
                INSERT ls_ref_formula INTO TABLE me->mt_ref_formulae.
              CATCH cx_root INTO ex.
                RAISE EXCEPTION TYPE zcx_excel
                  EXPORTING
                    previous = ex.
            ENDTRY.
          ENDIF.
*--------------------------------------------------------------------*
* End of insertion issue#284 - Copied formulae not
*--------------------------------------------------------------------*
        ENDIF.

        IF   lv_cell_value    IS NOT INITIAL
          OR lv_cell_formula  IS NOT INITIAL
          OR lv_style_guid    IS NOT INITIAL.
          zcl_excel_common=>convert_columnrow2column_a_row( EXPORTING
                                                              i_columnrow = ls_cell-r
                                                            IMPORTING
                                                              e_column    = lv_cell_column
                                                              e_row       = lv_cell_row ).
          io_worksheet->set_cell( ip_column     = lv_cell_column  " cell_elem Column
                                  ip_row        = lv_cell_row     " cell_elem row_elem
                                  ip_value      = lv_cell_value   " cell_elem Value
                                  ip_formula    = lv_cell_formula
                                  ip_data_type  = ls_cell-t
                                  ip_style      = lv_style_guid
                                  it_rtf        = lt_rtf ).
        ENDIF.
        lo_ixml_cell_elem ?= lo_ixml_iterator2->get_next( ).
      ENDWHILE.
      lo_ixml_row_elem ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.

*--------------------------------------------------------------------*
*#234 - column width not read correctly - begin of coding
*       reason - libre office doesn't use SPAN in row - definitions
*--------------------------------------------------------------------*
    IF lv_max_col = 0.
      lo_ixml_imension_elem = lo_ixml_worksheet->find_from_name_ns( name = 'dimension' uri = namespace-main ).
      IF lo_ixml_imension_elem IS BOUND.
        lv_dimension_range = lo_ixml_imension_elem->get_attribute( 'ref' ).
        IF lv_dimension_range CS ':'.
          REPLACE REGEX '\D+\d+:(\D+)\d+' IN lv_dimension_range WITH '$1'.  " Get max column
        ELSE.
          REPLACE REGEX '(\D+)\d+' IN lv_dimension_range WITH '$1'.  " Get max column
        ENDIF.
        lv_max_col = zcl_excel_common=>convert_column2int( lv_dimension_range ).
      ENDIF.
    ENDIF.
*--------------------------------------------------------------------*
*#234 - column width not read correctly - end of coding
*--------------------------------------------------------------------*

    "Get the customized column width
    lo_ixml_columns = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'col' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_columns->create_iterator( ).
    lo_ixml_column_elem ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml_column_elem IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_column_elem
                                   CHANGING
                                     cp_structure = ls_column ).
      lo_ixml_column_elem ?= lo_ixml_iterator->get_next( ).
      IF   ls_column-customwidth   = lc_xml_attr_true
        OR ls_column-customwidth   = lc_xml_attr_true_int
        OR ls_column-bestfit       = lc_xml_attr_true
        OR ls_column-bestfit       = lc_xml_attr_true_int
        OR ls_column-collapsed     = lc_xml_attr_true
        OR ls_column-collapsed     = lc_xml_attr_true_int
        OR ls_column-hidden        = lc_xml_attr_true
        OR ls_column-hidden        = lc_xml_attr_true_int
        OR ls_column-outlinelevel  > ''
        OR ls_column-style         > ''.
        lv_index = ls_column-min.
        WHILE lv_index <= ls_column-max AND lv_index <= lv_max_col.

          lv_column_alpha = zcl_excel_common=>convert_column2alpha( lv_index ).
          lo_column =  io_worksheet->get_column( lv_column_alpha ).

          IF   ls_column-customwidth = lc_xml_attr_true
            OR ls_column-customwidth = lc_xml_attr_true_int
            OR ls_column-width       IS NOT INITIAL.          "+#234
            lo_column->set_width( ls_column-width ).
          ENDIF.

          IF   ls_column-bestfit = lc_xml_attr_true
            OR ls_column-bestfit = lc_xml_attr_true_int.
            lo_column->set_auto_size( abap_true ).
          ENDIF.

          IF   ls_column-collapsed = lc_xml_attr_true
            OR ls_column-collapsed = lc_xml_attr_true_int.
            lo_column->set_collapsed( abap_true ).
          ENDIF.

          IF   ls_column-hidden = lc_xml_attr_true
            OR ls_column-hidden = lc_xml_attr_true_int.
            lo_column->set_visible( abap_false ).
          ENDIF.

          IF ls_column-outlinelevel > ''.
            CONDENSE ls_column-outlinelevel.
            lv_outline_level = ls_column-outlinelevel.
            IF lv_outline_level > 0.
              lo_column->set_outline_level( lv_outline_level ).
            ENDIF.
          ENDIF.

          IF ls_column-style > ''.
            lv_index_temp = ls_column-style + 1.
            READ TABLE styles INTO lo_excel_style INDEX lv_index_temp.
            DATA: dummy_zexcel_cell_style TYPE zexcel_cell_style.
            dummy_zexcel_cell_style = lo_excel_style->get_guid( ).
            lo_column->set_column_style_by_guid( dummy_zexcel_cell_style ).
          ENDIF.

          ADD 1 TO lv_index.
        ENDWHILE.
      ENDIF.

* issue #367 - hide columns from
      IF ls_column-max = zcl_excel_common=>c_excel_sheet_max_col.     " Max = very right column
        IF ls_column-hidden = 1     " all hidden
          AND ls_column-min > 0.
          io_worksheet->zif_excel_sheet_properties~hide_columns_from = zcl_excel_common=>convert_column2alpha( ls_column-min ).
        ELSEIF ls_column-style > ''.
          lv_index_temp = ls_column-style + 1.
          READ TABLE styles INTO lo_excel_style INDEX lv_index_temp.
          dummy_zexcel_cell_style = lo_excel_style->get_guid( ).
* Set style for remaining columns
          io_worksheet->zif_excel_sheet_properties~set_style( dummy_zexcel_cell_style ).
        ENDIF.
      ENDIF.


    ENDWHILE.

    "Now we need to get information from the sheetView node
    lo_ixml_sheetview_elem = lo_ixml_worksheet->find_from_name_ns( name = 'sheetView' uri = namespace-main ).
    fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_sheetview_elem CHANGING cp_structure = ls_sheetview ).
    IF ls_sheetview-showgridlines IS INITIAL OR
       ls_sheetview-showgridlines = lc_xml_attr_true OR
       ls_sheetview-showgridlines = lc_xml_attr_true_int.
      "If the attribute is not specified or set to true, we will show grid lines
      ls_sheetview-showgridlines = abap_true.
    ELSE.
      ls_sheetview-showgridlines = abap_false.
    ENDIF.
    io_worksheet->set_show_gridlines( ls_sheetview-showgridlines ).
    IF ls_sheetview-righttoleft = lc_xml_attr_true
        OR ls_sheetview-righttoleft = lc_xml_attr_true_int.
      io_worksheet->zif_excel_sheet_properties~set_right_to_left( abap_true ).
    ENDIF.
    io_worksheet->zif_excel_sheet_properties~zoomscale                 = ls_sheetview-zoomscale.
    io_worksheet->zif_excel_sheet_properties~zoomscale_normal          = ls_sheetview-zoomscalenormal.
    io_worksheet->zif_excel_sheet_properties~zoomscale_pagelayoutview  = ls_sheetview-zoomscalepagelayoutview.
    io_worksheet->zif_excel_sheet_properties~zoomscale_sheetlayoutview = ls_sheetview-zoomscalesheetlayoutview.
    IF ls_sheetview-topleftcell IS NOT INITIAL.
      io_worksheet->set_sheetview_top_left_cell( ls_sheetview-topleftcell ).
    ENDIF.

    "Add merge cell information
    lo_ixml_mergecells = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'mergeCell' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_mergecells->create_iterator( ).
    lo_ixml_mergecell_elem ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml_mergecell_elem IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_mergecell_elem
                                   CHANGING
                                     cp_structure = ls_mergecell ).
      zcl_excel_common=>convert_range2column_a_row( EXPORTING
                                                      i_range = ls_mergecell-ref
                                                    IMPORTING
                                                      e_column_start = lv_merge_column_start
                                                      e_column_end   = lv_merge_column_end
                                                      e_row_start    = lv_merge_row_start
                                                      e_row_end      = lv_merge_row_end ).
      lo_ixml_mergecell_elem ?= lo_ixml_iterator->get_next( ).
      io_worksheet->set_merge( EXPORTING
                                 ip_column_start = lv_merge_column_start
                                 ip_column_end   = lv_merge_column_end
                                 ip_row          = lv_merge_row_start
                                 ip_row_to       = lv_merge_row_end ).
    ENDWHILE.

    " read sheet format properties
    lo_ixml_sheetformatpr_elem = lo_ixml_worksheet->find_from_name_ns( name = 'sheetFormatPr' uri = namespace-main ).
    IF lo_ixml_sheetformatpr_elem IS NOT INITIAL.
      fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_sheetformatpr_elem CHANGING cp_structure = ls_sheetformatpr ).
      IF ls_sheetformatpr-customheight = '1'.
        lv_height = ls_sheetformatpr-defaultrowheight.
        lo_row = io_worksheet->get_default_row( ).
        lo_row->set_row_height( lv_height ).
      ENDIF.

      " TODO...  column
    ENDIF.

    " Read in page margins
    me->load_worksheet_pagemargins( EXPORTING
                                      io_ixml_worksheet = lo_ixml_worksheet
                                      io_worksheet      = io_worksheet ).

* FitToPage
    lo_ixml_sheetpr ?=  lo_ixml_worksheet->find_from_name_ns( name = 'pageSetUpPr' uri = namespace-main ).
    IF lo_ixml_sheetpr IS BOUND.

      lv_fit_to_page = lo_ixml_sheetpr->get_attribute_ns( 'fitToPage' ).
      IF lv_fit_to_page IS NOT INITIAL.
        io_worksheet->sheet_setup->fit_to_page = 'X'.
      ENDIF.
    ENDIF.
    " Read in page setup
    lo_ixml_pagesetup_elem = lo_ixml_worksheet->find_from_name_ns( name = 'pageSetup' uri = namespace-main ).
    IF lo_ixml_pagesetup_elem IS NOT INITIAL.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_pagesetup_elem
                                   CHANGING
                                     cp_structure = ls_pagesetup ).
      io_worksheet->sheet_setup->orientation = ls_pagesetup-orientation.
      io_worksheet->sheet_setup->scale = ls_pagesetup-scale.
      io_worksheet->sheet_setup->paper_size = ls_pagesetup-papersize.
      io_worksheet->sheet_setup->paper_height = ls_pagesetup-paperheight.
      io_worksheet->sheet_setup->paper_width = ls_pagesetup-paperwidth.
      IF io_worksheet->sheet_setup->fit_to_page = 'X'.
        IF ls_pagesetup-fittowidth IS NOT INITIAL.
          io_worksheet->sheet_setup->fit_to_width = ls_pagesetup-fittowidth.
        ELSE.
          io_worksheet->sheet_setup->fit_to_width = 1.  " Default if not given - Excel doesn't write this to xml
        ENDIF.
        IF ls_pagesetup-fittoheight IS NOT INITIAL.
          io_worksheet->sheet_setup->fit_to_height = ls_pagesetup-fittoheight.
        ELSE.
          io_worksheet->sheet_setup->fit_to_height = 1. " Default if not given - Excel doesn't write this to xml
        ENDIF.
      ENDIF.
    ENDIF.



    " Read header footer
    lo_ixml_headerfooter_elem = lo_ixml_worksheet->find_from_name_ns( name = 'headerFooter' uri = namespace-main ).
    IF lo_ixml_headerfooter_elem IS NOT INITIAL.
      fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_headerfooter_elem CHANGING cp_structure = ls_headerfooter ).
      io_worksheet->sheet_setup->diff_oddeven_headerfooter = ls_headerfooter-differentoddeven.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'oddFooter' uri = namespace-main ).
      IF lo_ixml_hf_value_elem IS NOT INITIAL.
        ls_odd_footer-left_value = lo_ixml_hf_value_elem->get_value( ).
      ENDIF.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'oddHeader' uri = namespace-main ).
      IF lo_ixml_hf_value_elem IS NOT INITIAL.
        ls_odd_header-left_value = lo_ixml_hf_value_elem->get_value( ).
      ENDIF.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'evenFooter' uri = namespace-main ).
      IF lo_ixml_hf_value_elem IS NOT INITIAL.
        ls_even_footer-left_value = lo_ixml_hf_value_elem->get_value( ).
      ENDIF.

      lo_ixml_hf_value_elem = lo_ixml_headerfooter_elem->find_from_name_ns( name = 'evenHeader' uri = namespace-main ).
      IF lo_ixml_hf_value_elem IS NOT INITIAL.
        ls_even_header-left_value = lo_ixml_hf_value_elem->get_value( ).
      ENDIF.

*        2do§1   Header/footer
      " TODO.. get the rest.

      io_worksheet->sheet_setup->set_header_footer( ip_odd_header   = ls_odd_header
                                                    ip_odd_footer   = ls_odd_footer
                                                    ip_even_header  = ls_even_header
                                                    ip_even_footer  = ls_even_footer ).

    ENDIF.

    " Read pane
    lo_ixml_pane_elem = lo_ixml_sheetview_elem->find_from_name_ns( name = 'pane' uri = namespace-main ).
    IF lo_ixml_pane_elem IS BOUND.
      fill_struct_from_attributes( EXPORTING ip_element = lo_ixml_pane_elem CHANGING cp_structure = ls_excel_pane ).
      lv_pane_cell_col = ls_excel_pane-xsplit.
      lv_pane_cell_row = ls_excel_pane-ysplit.
      IF    lv_pane_cell_col > 0
        AND lv_pane_cell_row > 0.
        io_worksheet->freeze_panes( ip_num_rows    = lv_pane_cell_row
                                    ip_num_columns = lv_pane_cell_col ).
      ELSEIF lv_pane_cell_row > 0.
        io_worksheet->freeze_panes( ip_num_rows    = lv_pane_cell_row ).
      ELSE.
        io_worksheet->freeze_panes( ip_num_columns = lv_pane_cell_col ).
      ENDIF.
      IF ls_excel_pane-topleftcell IS NOT INITIAL.
        io_worksheet->set_pane_top_left_cell( ls_excel_pane-topleftcell ).
      ENDIF.
    ENDIF.

    " Start fix 276 Read data validations
    lo_ixml_datavalidations = lo_ixml_worksheet->get_elements_by_tag_name_ns( name = 'dataValidation' uri = namespace-main ).
    lo_ixml_iterator = lo_ixml_datavalidations->create_iterator( ).
    lo_ixml_datavalidation_elem  ?= lo_ixml_iterator->get_next( ).
    WHILE lo_ixml_datavalidation_elem  IS BOUND.
      fill_struct_from_attributes( EXPORTING
                                     ip_element = lo_ixml_datavalidation_elem
                                   CHANGING
                                     cp_structure = ls_datavalidation ).
      CLEAR lo_ixml_formula_elem.
      lo_ixml_formula_elem = lo_ixml_datavalidation_elem->find_from_name_ns( name = 'formula1' uri = namespace-main ).
      IF lo_ixml_formula_elem IS BOUND.
        ls_datavalidation-formula1 = lo_ixml_formula_elem->get_value( ).
      ENDIF.
      CLEAR lo_ixml_formula_elem.
      lo_ixml_formula_elem = lo_ixml_datavalidation_elem->find_from_name_ns( name = 'formula2' uri = namespace-main ).
      IF lo_ixml_formula_elem IS BOUND.
        ls_datavalidation-formula2 = lo_ixml_formula_elem->get_value( ).
      ENDIF.
      SPLIT ls_datavalidation-sqref AT space INTO TABLE lt_datavalidation_range.
      LOOP AT lt_datavalidation_range INTO lv_datavalidation_range.
        zcl_excel_common=>convert_range2column_a_row( EXPORTING
                                                        i_range = lv_datavalidation_range
                                                      IMPORTING
                                                        e_column_start = ls_datavalidation-cell_column
                                                        e_column_end   = ls_datavalidation-cell_column_to
                                                        e_row_start    = ls_datavalidation-cell_row
                                                        e_row_end      = ls_datavalidation-cell_row_to ).
        lo_data_validation                   = io_worksheet->add_new_data_validation( ).
        lo_data_validation->type             = ls_datavalidation-type.
        lo_data_validation->allowblank       = ls_datavalidation-allowblank.
        IF ls_datavalidation-showinputmessage IS INITIAL.
          lo_data_validation->showinputmessage = abap_false.
        ELSE.
          lo_data_validation->showinputmessage = abap_true.
        ENDIF.
        IF ls_datavalidation-showerrormessage IS INITIAL.
          lo_data_validation->showerrormessage = abap_false.
        ELSE.
          lo_data_validation->showerrormessage = abap_true.
        ENDIF.
        IF ls_datavalidation-showdropdown IS INITIAL.
          lo_data_validation->showdropdown = abap_false.
        ELSE.
          lo_data_validation->showdropdown = abap_true.
        ENDIF.
        lo_data_validation->operator         = ls_datavalidation-operator.
        lo_data_validation->formula1         = ls_datavalidation-formula1.
        lo_data_validation->formula2         = ls_datavalidation-formula2.
        lo_data_validation->prompttitle      = ls_datavalidation-prompttitle.
        lo_data_validation->prompt           = ls_datavalidation-prompt.
        lo_data_validation->errortitle       = ls_datavalidation-errortitle.
        lo_data_validation->error            = ls_datavalidation-error.
        lo_data_validation->errorstyle       = ls_datavalidation-errorstyle.
        lo_data_validation->cell_row         = ls_datavalidation-cell_row.
        lo_data_validation->cell_row_to      = ls_datavalidation-cell_row_to.
        lo_data_validation->cell_column      = ls_datavalidation-cell_column.
        lo_data_validation->cell_column_to   = ls_datavalidation-cell_column_to.
      ENDLOOP.
      lo_ixml_datavalidation_elem ?= lo_ixml_iterator->get_next( ).
    ENDWHILE.
    " End fix 276 Read data validations

    " Read hyperlinks
    TRY.
        me->load_worksheet_hyperlinks( io_ixml_worksheet      = lo_ixml_worksheet
                                       io_worksheet           = io_worksheet
                                       it_external_hyperlinks = lt_external_hyperlinks ).
      CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    ENDTRY.

    TRY.
        me->fill_row_outlines( io_worksheet           = io_worksheet ).
      CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    ENDTRY.

    " Issue #366 - conditional formatting
    TRY.
        me->load_worksheet_cond_format( io_ixml_worksheet      = lo_ixml_worksheet
                                        io_worksheet           = io_worksheet ).
      CATCH zcx_excel. " Ignore Hyperlink reading errors - pass everything we were able to identify
    ENDTRY.

    " Issue #377 - pagebreaks
    TRY.
        me->load_worksheet_pagebreaks( io_ixml_worksheet      = lo_ixml_worksheet
                                       io_worksheet           = io_worksheet ).
      CATCH zcx_excel. " Ignore pagebreak reading errors - pass everything we were able to identify
    ENDTRY.

    TRY.
        me->load_worksheet_autofilter( io_ixml_worksheet      = lo_ixml_worksheet
                                       io_worksheet           = io_worksheet ).
      CATCH zcx_excel. " Ignore autofilter reading errors - pass everything we were able to identify
    ENDTRY.

    TRY.
        me->load_worksheet_ignored_errors( io_ixml_worksheet      = lo_ixml_worksheet
                                           io_worksheet           = io_worksheet ).
      CATCH zcx_excel. " Ignore "ignoredErrors" reading errors - pass everything we were able to identify
    ENDTRY.

  ENDMETHOD.