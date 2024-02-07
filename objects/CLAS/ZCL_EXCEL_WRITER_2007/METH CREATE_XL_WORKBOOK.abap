  METHOD create_xl_workbook.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-07
*              - ...
* changes: aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                             2012-12-01
* changes:  correction of pointer to localSheetId
*--------------------------------------------------------------------*

** Constant node name
    DATA: lc_xml_node_workbook           TYPE string VALUE 'workbook',
          lc_xml_node_fileversion        TYPE string VALUE 'fileVersion',
          lc_xml_node_workbookpr         TYPE string VALUE 'workbookPr',
          lc_xml_node_bookviews          TYPE string VALUE 'bookViews',
          lc_xml_node_workbookview       TYPE string VALUE 'workbookView',
          lc_xml_node_sheets             TYPE string VALUE 'sheets',
          lc_xml_node_sheet              TYPE string VALUE 'sheet',
          lc_xml_node_calcpr             TYPE string VALUE 'calcPr',
          lc_xml_node_workbookprotection TYPE string VALUE 'workbookProtection',
          lc_xml_node_definednames       TYPE string VALUE 'definedNames',
          lc_xml_node_definedname        TYPE string VALUE 'definedName',
          " Node attributes
          lc_xml_attr_appname            TYPE string VALUE 'appName',
          lc_xml_attr_lastedited         TYPE string VALUE 'lastEdited',
          lc_xml_attr_lowestedited       TYPE string VALUE 'lowestEdited',
          lc_xml_attr_rupbuild           TYPE string VALUE 'rupBuild',
          lc_xml_attr_xwindow            TYPE string VALUE 'xWindow',
          lc_xml_attr_ywindow            TYPE string VALUE 'yWindow',
          lc_xml_attr_windowwidth        TYPE string VALUE 'windowWidth',
          lc_xml_attr_windowheight       TYPE string VALUE 'windowHeight',
          lc_xml_attr_activetab          TYPE string VALUE 'activeTab',
          lc_xml_attr_name               TYPE string VALUE 'name',
          lc_xml_attr_sheetid            TYPE string VALUE 'sheetId',
          lc_xml_attr_state              TYPE string VALUE 'state',
          lc_xml_attr_id                 TYPE string VALUE 'id',
          lc_xml_attr_calcid             TYPE string VALUE 'calcId',
          lc_xml_attr_lockrevision       TYPE string VALUE 'lockRevision',
          lc_xml_attr_lockstructure      TYPE string VALUE 'lockStructure',
          lc_xml_attr_lockwindows        TYPE string VALUE 'lockWindows',
          lc_xml_attr_revisionspassword  TYPE string VALUE 'revisionsPassword',
          lc_xml_attr_workbookpassword   TYPE string VALUE 'workbookPassword',
          lc_xml_attr_hidden             TYPE string VALUE 'hidden',
          lc_xml_attr_localsheetid       TYPE string VALUE 'localSheetId',
          " Node namespace
          lc_r_ns                        TYPE string VALUE 'r',
          lc_xml_node_ns                 TYPE string VALUE 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          lc_xml_node_r_ns               TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          " Node id
          lc_xml_node_ridx_id            TYPE string VALUE 'rId#'.

    DATA: lo_document       TYPE REF TO if_ixml_document,
          lo_element_root   TYPE REF TO if_ixml_element,
          lo_element        TYPE REF TO if_ixml_element,
          lo_element_range  TYPE REF TO if_ixml_element,
          lo_sub_element    TYPE REF TO if_ixml_element,
          lo_iterator       TYPE REF TO zcl_excel_collection_iterator,
          lo_iterator_range TYPE REF TO zcl_excel_collection_iterator,
          lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          lo_range          TYPE REF TO zcl_excel_range,
          lo_autofilters    TYPE REF TO zcl_excel_autofilters,
          lo_autofilter     TYPE REF TO zcl_excel_autofilter.

    DATA: lv_xml_node_ridx_id TYPE string,
          lv_value            TYPE string,
          lv_syindex          TYPE string,
          lv_active_sheet     TYPE zexcel_active_worksheet.

**********************************************************************
* STEP 1: Create [Content_Types].xml into the root of the ZIP
    lo_document = create_xml_document( ).

**********************************************************************
* STEP 3: Create main node
    lo_element_root  = lo_document->create_simple_element( name   = lc_xml_node_workbook
                                                           parent = lo_document ).
    lo_element_root->set_attribute_ns( name  = 'xmlns'
                                       value = lc_xml_node_ns ).
    lo_element_root->set_attribute_ns( name  = 'xmlns:r'
                                       value = lc_xml_node_r_ns ).

**********************************************************************
* STEP 4: Create subnode
    " fileVersion node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_fileversion
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_appname
                                  value = 'xl' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_lastedited
                                  value = '4' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_lowestedited
                                  value = '4' ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_rupbuild
                                  value = '4506' ).
    lo_element_root->append_child( new_child = lo_element ).

    " fileVersion node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_workbookpr
                                                     parent = lo_document ).
    lo_element_root->append_child( new_child = lo_element ).

    " workbookProtection node
    IF me->excel->zif_excel_book_protection~protected EQ abap_true.
      lo_element = lo_document->create_simple_element( name   = lc_xml_node_workbookprotection
                                                       parent = lo_document ).
      lv_value = me->excel->zif_excel_book_protection~workbookpassword.
      IF lv_value IS NOT INITIAL.
        lo_element->set_attribute_ns( name  = lc_xml_attr_workbookpassword
                                      value = lv_value ).
      ENDIF.
      lv_value = me->excel->zif_excel_book_protection~revisionspassword.
      IF lv_value IS NOT INITIAL.
        lo_element->set_attribute_ns( name  = lc_xml_attr_revisionspassword
                                      value = lv_value ).
      ENDIF.
      lv_value = me->excel->zif_excel_book_protection~lockrevision.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockrevision
                                    value = lv_value ).
      lv_value = me->excel->zif_excel_book_protection~lockstructure.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockstructure
                                    value = lv_value ).
      lv_value = me->excel->zif_excel_book_protection~lockwindows.
      CONDENSE lv_value NO-GAPS.
      lo_element->set_attribute_ns( name  = lc_xml_attr_lockwindows
                                    value = lv_value ).
      lo_element_root->append_child( new_child = lo_element ).
    ENDIF.

    " bookviews node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_bookviews
                                                     parent = lo_document ).
    " bookview node
    lo_sub_element = lo_document->create_simple_element( name   = lc_xml_node_workbookview
                                                         parent = lo_document ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_xwindow
                                      value = '120' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_ywindow
                                      value = '120' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_windowwidth
                                      value = '19035' ).
    lo_sub_element->set_attribute_ns( name  = lc_xml_attr_windowheight
                                      value = '8445' ).
    " Set Active Sheet
    lv_active_sheet = excel->get_active_sheet_index( ).
* issue #365 - test if sheet exists - otherwise set active worksheet to 1
    lo_worksheet = excel->get_worksheet_by_index( lv_active_sheet ).
    IF lo_worksheet IS NOT BOUND.
      lv_active_sheet = 1.
      excel->set_active_sheet_index( lv_active_sheet ).
    ENDIF.
    IF lv_active_sheet > 1.
      lv_active_sheet = lv_active_sheet - 1.
      lv_value = lv_active_sheet.
      CONDENSE lv_value.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_activetab
                                        value = lv_value ).
    ENDIF.
    lo_element->append_child( new_child = lo_sub_element )." bookview node
    lo_element_root->append_child( new_child = lo_element )." bookviews node

    " sheets node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_sheets
                                                     parent = lo_document ).
    lo_iterator = excel->get_worksheets_iterator( ).

    " ranges node
    lo_element_range = lo_document->create_simple_element( name   = lc_xml_node_definednames " issue 163 +
                                                           parent = lo_document ).           " issue 163 +

    WHILE lo_iterator->has_next( ) EQ abap_true.
      " sheet node
      lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_sheet
                                                              parent = lo_document ).
      lo_worksheet ?= lo_iterator->get_next( ).
      lv_syindex = sy-index.                                                                  " question by Stefan SchmÃ¶cker 2012-12-02:  sy-index seems to do the job - but is it proven to work or purely coincedence
      lv_value = lo_worksheet->get_title( ).
      SHIFT lv_syindex RIGHT DELETING TRAILING space.
      SHIFT lv_syindex LEFT DELETING LEADING space.
      lv_xml_node_ridx_id = lc_xml_node_ridx_id.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xml_node_ridx_id WITH lv_syindex.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                        value = lv_value ).
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_sheetid
                                        value = lv_syindex ).
      IF lo_worksheet->zif_excel_sheet_properties~hidden EQ zif_excel_sheet_properties=>c_hidden.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_state
                                          value = 'hidden' ).
      ELSEIF lo_worksheet->zif_excel_sheet_properties~hidden EQ zif_excel_sheet_properties=>c_veryhidden.
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_state
                                          value = 'veryHidden' ).
      ENDIF.
      lo_sub_element->set_attribute_ns( name    = lc_xml_attr_id
                                        prefix  = lc_r_ns
                                        value   = lv_xml_node_ridx_id ).
      lo_element->append_child( new_child = lo_sub_element ). " sheet node

      " issue 163 >>>
      lo_iterator_range = lo_worksheet->get_ranges_iterator( ).

*--------------------------------------------------------------------*
* Defined names sheetlocal:  Ranges, Repeat rows and columns
*--------------------------------------------------------------------*
      WHILE lo_iterator_range->has_next( ) EQ abap_true.
        " range node
        lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                                parent = lo_document ).
        lo_range ?= lo_iterator_range->get_next( ).
        lv_value = lo_range->name.

        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                          value = lv_value ).

*      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid           "del #235 Repeat rows/cols - EXCEL starts couting from zero
*                                        value = lv_xml_node_ridx_id ).             "del #235 Repeat rows/cols - and needs absolute referencing to localSheetId
        lv_value   = lv_syindex - 1.                                                  "ins #235 Repeat rows/cols
        CONDENSE lv_value NO-GAPS.                                                    "ins #235 Repeat rows/cols
        lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid
                                          value = lv_value ).

        lv_value = lo_range->get_value( ).
        lo_sub_element->set_value( value = lv_value ).
        lo_element_range->append_child( new_child = lo_sub_element ). " range node

      ENDWHILE.
      " issue 163 <<<

    ENDWHILE.
    lo_element_root->append_child( new_child = lo_element )." sheets node


*--------------------------------------------------------------------*
* Defined names workbookgolbal:  Ranges
*--------------------------------------------------------------------*
*  " ranges node
*  lo_element = lo_document->create_simple_element( name   = lc_xml_node_definednames " issue 163 -
*                                                   parent = lo_document ).           " issue 163 -
    lo_iterator = excel->get_ranges_iterator( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      " range node
      lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                              parent = lo_document ).
      lo_range ?= lo_iterator->get_next( ).
      lv_value = lo_range->name.
      lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                        value = lv_value ).
      lv_value = lo_range->get_value( ).
      lo_sub_element->set_value( value = lv_value ).
      lo_element_range->append_child( new_child = lo_sub_element ). " range node

    ENDWHILE.

*--------------------------------------------------------------------*
* Defined names - Autofilters ( also sheetlocal )
*--------------------------------------------------------------------*
    lo_autofilters = excel->get_autofilters_reference( ).
    IF lo_autofilters->is_empty( ) = abap_false.
      lo_iterator = excel->get_worksheets_iterator( ).
      WHILE lo_iterator->has_next( ) EQ abap_true.

        lo_worksheet ?= lo_iterator->get_next( ).
        lv_syindex = sy-index - 1 .
        lo_autofilter = lo_autofilters->get( io_worksheet = lo_worksheet ).
        IF lo_autofilter IS BOUND.
          lo_sub_element = lo_document->create_simple_element_ns( name   = lc_xml_node_definedname
                                                                  parent = lo_document ).
          lv_value = lo_autofilters->c_autofilter.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_name
                                            value = lv_value ).
          lv_value = lv_syindex.
          CONDENSE lv_value NO-GAPS.
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_localsheetid
                                            value = lv_value ).
          lv_value = '1'. " Always hidden
          lo_sub_element->set_attribute_ns( name  = lc_xml_attr_hidden
                                            value = lv_value ).
          lv_value = lo_autofilter->get_filter_reference( ).
          lo_sub_element->set_value( value = lv_value ).
          lo_element_range->append_child( new_child = lo_sub_element ). " range node
        ENDIF.

      ENDWHILE.
    ENDIF.
    lo_element_root->append_child( new_child = lo_element_range ).                      " ranges node


    " calcPr node
    lo_element = lo_document->create_simple_element( name   = lc_xml_node_calcpr
                                                     parent = lo_document ).
    lo_element->set_attribute_ns( name  = lc_xml_attr_calcid
                                  value = '125725' ).
    lo_element_root->append_child( new_child = lo_element ).

**********************************************************************
* STEP 5: Create xstring stream
    ep_content = render_xml_document( lo_document ).

  ENDMETHOD.