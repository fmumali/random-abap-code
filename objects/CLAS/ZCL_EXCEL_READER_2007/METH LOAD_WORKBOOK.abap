  METHOD load_workbook.
*--------------------------------------------------------------------*
* ToDos:
*        2do§1   Move macro-reading from zcl_excel_reader_xlsm to this class
*                autodetect existance of macro/vba content
*                Allow inputparameter to explicitly tell reader to ignore vba-content
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-10
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          removing unused variables
*          adding me-> where possible
*          renaming variables to indicate what they are used for
*          adding comments to explain what we are trying to achieve
*          renaming i/o parameters:  previous input-parameter ip_path  holds a (full) filename and not a path   --> rename to iv_workbook_full_filename
*                                                             ip_excel renamed while being at it                --> rename to io_excel
*--------------------------------------------------------------------*
* issue #232   - Read worksheetstate hidden/veryHidden
*              - Stefan Schmoecker,                          2012-11-11
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns
*           - Stefan Schmoecker,                             2012-12-02
* changes:    correction in named ranges to correctly attach
*             sheetlocal names/ranges to the correct sheet
*--------------------------------------------------------------------*
* issue#284 - Copied formulae ignored when reading excelfile
*           - Stefan Schmoecker,                             2013-08-02
* changes:    initialize area to hold referenced formulaedata
*             after all worksheets have been read resolve formuae
*--------------------------------------------------------------------*

    CONSTANTS: lcv_shared_strings             TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
               lcv_worksheet                  TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
               lcv_styles                     TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
               lcv_vba_project                TYPE string VALUE 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', "#EC NEEDED     for future incorporation of XLSM-reader
               lcv_theme                      TYPE string VALUE 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
*--------------------------------------------------------------------*
* #232: Read worksheetstate hidden/veryHidden - begin data declarations
*--------------------------------------------------------------------*
               lcv_worksheet_state_hidden     TYPE string VALUE 'hidden',
               lcv_worksheet_state_veryhidden TYPE string VALUE 'veryHidden'.
*--------------------------------------------------------------------*
* #232: Read worksheetstate hidden/veryHidden - end data declarations
*--------------------------------------------------------------------*

    DATA:
      lv_path                    TYPE string,
      lv_filename                TYPE chkfile,
      lv_full_filename           TYPE string,

      lo_rels_workbook           TYPE REF TO if_ixml_document,
      lt_worksheets              TYPE STANDARD TABLE OF t_relationship WITH NON-UNIQUE DEFAULT KEY,
      lo_workbook                TYPE REF TO if_ixml_document,
      lv_workbook_index          TYPE i,
      lv_worksheet_path          TYPE string,
      ls_sheet                   TYPE t_sheet,

      lo_node                    TYPE REF TO if_ixml_element,
      ls_relationship            TYPE t_relationship,
      lo_worksheet               TYPE REF TO zcl_excel_worksheet,
      lo_range                   TYPE REF TO zcl_excel_range,
      lv_worksheet_title         TYPE zexcel_sheet_title,
      lv_tabix                   TYPE i,            " #235 - repeat rows/cols.  Needed to link defined name to correct worksheet

      ls_range                   TYPE t_range,
      lv_range_value             TYPE zexcel_range_value,
      lv_position_temp           TYPE i,
*--------------------------------------------------------------------*
* #229: Set active worksheet - begin data declarations
*--------------------------------------------------------------------*
      lv_active_sheet_string     TYPE string,
      lv_zexcel_active_worksheet TYPE zexcel_active_worksheet,
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns  - added autofilter support while changing this section
      lo_autofilter              TYPE REF TO zcl_excel_autofilter,
      ls_area                    TYPE zexcel_s_autofilter_area,
      lv_col_start_alpha         TYPE zexcel_cell_column_alpha,
      lv_col_end_alpha           TYPE zexcel_cell_column_alpha,
      lv_row_start               TYPE zexcel_cell_row,
      lv_row_end                 TYPE zexcel_cell_row,
      lv_regex                   TYPE string,
      lv_range_value_1           TYPE zexcel_range_value,
      lv_range_value_2           TYPE zexcel_range_value.
*--------------------------------------------------------------------*
* #229: Set active worksheet - end data declarations
*--------------------------------------------------------------------*
    FIELD-SYMBOLS: <worksheet> TYPE t_relationship.


*--------------------------------------------------------------------*

* §1  Get the position of files related to this workbook
*         Usually this will be <root>/xl/workbook.xml
*         Thus the workbookroot will be <root>/xl/
*         The position of all related files will be given in file
*         <workbookroot>/_rels/<workbookfilename>.rels and their positions
*         be be given relative to the workbookroot

*     Following is an example how this file could be set up

*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
*            <Relationship Target="styles.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Id="rId6"/>
*            <Relationship Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Id="rId5"/>
*            <Relationship Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId1"/>
*            <Relationship Target="worksheets/sheet2.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId2"/>
*            <Relationship Target="worksheets/sheet3.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId3"/>
*            <Relationship Target="worksheets/sheet4.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Id="rId4"/>
*            <Relationship Target="sharedStrings.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Id="rId7"/>
*        </Relationships>
*
* §2  Load data that is relevant to the complete workbook
*     Currently supported is:
*   §2.1    Shared strings  - This holds all strings that are used in all worksheets
*   §2.2    Styles          - This holds all styles that are used in all worksheets
*   §2.3    Worksheets      - For each worksheet in the workbook one entry appears here to point to the file that holds the content of this worksheet
*   §2.4    [Themes]                - not supported
*   §2.5    [VBA (Macro)]           - supported in class zcl_excel_reader_xlsm but should be moved here and autodetect
*   ...
*
* §3  Some information is held in the workbookfile as well
*   §3.1    Names and order of of worksheets
*   §3.2    Active worksheet
*   §3.3    Defined names
*   ...
*     Following is an example how this file could be set up

*        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
*        <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
*            <fileVersion rupBuild="4506" lowestEdited="4" lastEdited="4" appName="xl"/>
*            <workbookPr defaultThemeVersion="124226"/>
*            <bookViews>
*                <workbookView activeTab="1" windowHeight="8445" windowWidth="19035" yWindow="120" xWindow="120"/>
*            </bookViews>
*            <sheets>
*                <sheet r:id="rId1" sheetId="1" name="Sheet1"/>
*                <sheet r:id="rId2" sheetId="2" name="Sheet2"/>
*                <sheet r:id="rId3" sheetId="3" name="Sheet3" state="hidden"/>
*                <sheet r:id="rId4" sheetId="4" name="Sheet4"/>
*            </sheets>
*            <definedNames/>
*            <calcPr calcId="125725"/>
*        </workbook>
*--------------------------------------------------------------------*

    CLEAR me->mt_ref_formulae.                                                                              " ins issue#284

*--------------------------------------------------------------------*
* §1  Get the position of files related to this workbook
*     Entry into this method is with the filename of the workbook
*--------------------------------------------------------------------*
    CALL FUNCTION 'TRINT_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = iv_workbook_full_filename
      IMPORTING
        stripped_name = lv_filename
        file_path     = lv_path.

    CONCATENATE lv_path '_rels/' lv_filename '.rels'
        INTO lv_full_filename.
    lo_rels_workbook = me->get_ixml_from_zip_archive( lv_full_filename ).

    lo_node ?= lo_rels_workbook->find_from_name_ns( name = 'Relationship' uri = namespace-relationships ). "#EC NOTEXT
    WHILE lo_node IS BOUND.

      me->fill_struct_from_attributes( EXPORTING ip_element = lo_node CHANGING cp_structure = ls_relationship ).

      CASE ls_relationship-type.

*--------------------------------------------------------------------*
*   §2.1    Shared strings  - This holds all strings that are used in all worksheets
*--------------------------------------------------------------------*
        WHEN lcv_shared_strings.
          CONCATENATE lv_path ls_relationship-target
              INTO lv_full_filename.
          me->load_shared_strings( lv_full_filename ).

*--------------------------------------------------------------------*
*   §2.3    Worksheets
*           For each worksheet in the workbook one entry appears here to point to the file that holds the content of this worksheet
*           Shared strings and styles have to be present before we can start with creating the worksheets
*           thus we only store this information for use when parsing the workbookfile for sheetinformations
*--------------------------------------------------------------------*
        WHEN lcv_worksheet.
          APPEND ls_relationship TO lt_worksheets.

*--------------------------------------------------------------------*
*   §2.2    Styles           - This holds the styles that are used in all worksheets
*--------------------------------------------------------------------*
        WHEN lcv_styles.
          CONCATENATE lv_path ls_relationship-target
              INTO lv_full_filename.
          me->load_styles( ip_path  = lv_full_filename
                           ip_excel = io_excel ).
          me->load_dxf_styles( iv_path  = lv_full_filename
                               io_excel = io_excel ).
        WHEN lcv_theme.
          CONCATENATE lv_path ls_relationship-target
              INTO lv_full_filename.
          me->load_theme(
                  EXPORTING
                    iv_path  = lv_full_filename
                    ip_excel = io_excel   " Excel creator
                    ).
        WHEN OTHERS.

      ENDCASE.

      lo_node ?= lo_node->get_next( ).

    ENDWHILE.

*--------------------------------------------------------------------*
* §3  Some information held in the workbookfile
*--------------------------------------------------------------------*
    lo_workbook = me->get_ixml_from_zip_archive( iv_workbook_full_filename ).

*--------------------------------------------------------------------*
*   §3.1    Names and order of of worksheets
*--------------------------------------------------------------------*
    lo_node           ?= lo_workbook->find_from_name_ns( name = 'sheet' uri = namespace-main ).
    lv_workbook_index  = 1.
    WHILE lo_node IS BOUND.

      me->fill_struct_from_attributes( EXPORTING
                                         ip_element   = lo_node
                                       CHANGING
                                         cp_structure = ls_sheet ).
*--------------------------------------------------------------------*
*       Create new worksheet in workbook with correct name
*--------------------------------------------------------------------*
      lv_worksheet_title = ls_sheet-name.
      IF lv_workbook_index = 1.                                               " First sheet has been added automatically by creating io_excel
        lo_worksheet = io_excel->get_active_worksheet( ).
        lo_worksheet->set_title( lv_worksheet_title ).
      ELSE.
        lo_worksheet = io_excel->add_new_worksheet( lv_worksheet_title ).
      ENDIF.
*--------------------------------------------------------------------*
* #232   - Read worksheetstate hidden/veryHidden - begin of coding
*       Set status hidden if necessary
*--------------------------------------------------------------------*
      CASE ls_sheet-state.

        WHEN lcv_worksheet_state_hidden.
          lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_hidden.

        WHEN lcv_worksheet_state_veryhidden.
          lo_worksheet->zif_excel_sheet_properties~hidden = zif_excel_sheet_properties=>c_veryhidden.

      ENDCASE.
*--------------------------------------------------------------------*
* #232   - Read worksheetstate hidden/veryHidden - end of coding
*--------------------------------------------------------------------*
*--------------------------------------------------------------------*
*       Load worksheetdata
*--------------------------------------------------------------------*
      READ TABLE lt_worksheets ASSIGNING <worksheet> WITH KEY id = ls_sheet-id.
      IF sy-subrc = 0.
        <worksheet>-sheetid = ls_sheet-sheetid.                                "ins #235 - repeat rows/cols - needed to identify correct sheet
        <worksheet>-localsheetid = |{ lv_workbook_index - 1 }|.
        CONCATENATE lv_path <worksheet>-target
            INTO lv_worksheet_path.
        me->load_worksheet( ip_path      = lv_worksheet_path
                            io_worksheet = lo_worksheet ).
        <worksheet>-worksheet = lo_worksheet.
      ENDIF.

      lo_node ?= lo_node->get_next( ).
      ADD 1 TO lv_workbook_index.

    ENDWHILE.
    SORT lt_worksheets BY sheetid.                                              " needed for localSheetid -referencing

*--------------------------------------------------------------------*
*   #284: Set active worksheet - Resolve referenced formulae to
*                                explicit formulae those cells
*--------------------------------------------------------------------*
    me->resolve_referenced_formulae( ).
    " ins issue#284
*--------------------------------------------------------------------*
*   #229: Set active worksheet - begin coding
*   §3.2    Active worksheet
*--------------------------------------------------------------------*
    lv_zexcel_active_worksheet = 1.                                 " First sheet = active sheet if nothing else specified.
    lo_node ?=  lo_workbook->find_from_name_ns( name = 'workbookView' uri = namespace-main ).
    IF lo_node IS BOUND.
      lv_active_sheet_string = lo_node->get_attribute( 'activeTab' ).
      TRY.
          lv_zexcel_active_worksheet = lv_active_sheet_string + 1.  " EXCEL numbers the sheets from 0 onwards --> index into worksheettable is increased by one
        CATCH cx_sy_conversion_error. "#EC NO_HANDLER    - error here --> just use the default 1st sheet
      ENDTRY.
    ENDIF.
    io_excel->set_active_sheet_index( lv_zexcel_active_worksheet ).
*--------------------------------------------------------------------*
* #229: Set active worksheet - end coding
*--------------------------------------------------------------------*


*--------------------------------------------------------------------*
*   §3.3    Defined names
*           So far I have encountered these
*             - named ranges      - sheetlocal
*             - named ranges      - workbookglobal
*             - autofilters       - sheetlocal  ( special range )
*             - repeat rows/cols  - sheetlocal ( special range )
*
*--------------------------------------------------------------------*
    lo_node ?=  lo_workbook->find_from_name_ns( name = 'definedName' uri = namespace-main ).
    WHILE lo_node IS BOUND.

      CLEAR lo_range.                                                                                       "ins issue #235 - repeat rows/cols
      me->fill_struct_from_attributes(  EXPORTING
                                        ip_element   =  lo_node
                                       CHANGING
                                         cp_structure = ls_range ).
      lv_range_value = lo_node->get_value( ).

      IF ls_range-localsheetid IS NOT INITIAL.                                                              " issue #163+
*      READ TABLE lt_worksheets ASSIGNING <worksheet> WITH KEY id = ls_range-localsheetid.                "del issue #235 - repeat rows/cols " issue #163+
*        lo_range = <worksheet>-worksheet->add_new_range( ).                                              "del issue #235 - repeat rows/cols " issue #163+
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns - begin
*--------------------------------------------------------------------*
        READ TABLE lt_worksheets ASSIGNING <worksheet> WITH KEY localsheetid = ls_range-localsheetid.
        IF sy-subrc = 0.
          CASE ls_range-name.

*--------------------------------------------------------------------*
* insert autofilters
*--------------------------------------------------------------------*
            WHEN zcl_excel_autofilters=>c_autofilter.
              " begin Dennis Schaaf
              TRY.
                  zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range        = lv_range_value
                                                                IMPORTING e_column_start = lv_col_start_alpha
                                                                          e_column_end   = lv_col_end_alpha
                                                                          e_row_start    = ls_area-row_start
                                                                          e_row_end      = ls_area-row_end ).
                  ls_area-col_start = zcl_excel_common=>convert_column2int( lv_col_start_alpha ).
                  ls_area-col_end   = zcl_excel_common=>convert_column2int( lv_col_end_alpha ).
                  lo_autofilter = io_excel->add_new_autofilter( io_sheet = <worksheet>-worksheet ) .
                  lo_autofilter->set_filter_area( is_area = ls_area ).
                CATCH zcx_excel.
                  " we expected a range but it was not usable, so just ignore it
              ENDTRY.
              " end Dennis Schaaf

*--------------------------------------------------------------------*
* repeat print rows/columns
*--------------------------------------------------------------------*
            WHEN zif_excel_sheet_printsettings=>gcv_print_title_name.
              lo_range = <worksheet>-worksheet->add_new_range( ).
              lo_range->name = zif_excel_sheet_printsettings=>gcv_print_title_name.
*--------------------------------------------------------------------*
* This might be a temporary solution.  Maybe ranges get be reworked
* to support areas consisting of multiple rectangles
* But for now just split the range into row and columnpart
*--------------------------------------------------------------------*
              CLEAR:lv_range_value_1,
                    lv_range_value_2.
              IF lv_range_value IS INITIAL.
* Empty --> nothing to do
              ELSE.
                IF lv_range_value(1) = `'`.  " Escaped
                  lv_regex = `^('[^']*')+![^,]*,`.
                ELSE.
                  lv_regex = `^[^!]*![^,]*,`.
                ENDIF.
* Split into two ranges if necessary
                FIND REGEX lv_regex IN lv_range_value MATCH LENGTH lv_position_temp.
                IF sy-subrc = 0 AND lv_position_temp > 0.
                  lv_range_value_2 = lv_range_value+lv_position_temp.
                  SUBTRACT 1 FROM lv_position_temp.
                  lv_range_value_1 = lv_range_value(lv_position_temp).
                ELSE.
                  lv_range_value_1 = lv_range_value.
                ENDIF.
              ENDIF.
* 1st range
              zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range            = lv_range_value_1
                                                                      i_allow_1dim_range = abap_true
                                                            IMPORTING e_column_start     = lv_col_start_alpha
                                                                      e_column_end       = lv_col_end_alpha
                                                                      e_row_start        = lv_row_start
                                                                      e_row_end          = lv_row_end ).
              IF lv_col_start_alpha IS NOT INITIAL.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = lv_col_start_alpha
                                                                                      iv_columns_to   = lv_col_end_alpha ).
              ENDIF.
              IF lv_row_start IS NOT INITIAL.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = lv_row_start
                                                                                   iv_rows_to   = lv_row_end ).
              ENDIF.

* 2nd range
              zcl_excel_common=>convert_range2column_a_row( EXPORTING i_range            = lv_range_value_2
                                                                      i_allow_1dim_range = abap_true
                                                            IMPORTING e_column_start     = lv_col_start_alpha
                                                                      e_column_end       = lv_col_end_alpha
                                                                      e_row_start        = lv_row_start
                                                                      e_row_end          = lv_row_end ).
              IF lv_col_start_alpha IS NOT INITIAL.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_columns( iv_columns_from = lv_col_start_alpha
                                                                                      iv_columns_to   = lv_col_end_alpha ).
              ENDIF.
              IF lv_row_start IS NOT INITIAL.
                <worksheet>-worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows( iv_rows_from = lv_row_start
                                                                                   iv_rows_to   = lv_row_end ).
              ENDIF.

            WHEN OTHERS.

          ENDCASE.
        ENDIF.
*--------------------------------------------------------------------*
* issue#235 - repeat rows/columns - end
*--------------------------------------------------------------------*
      ELSE.                                                                                                 " issue #163+
        lo_range = io_excel->add_new_range( ).                                                              " issue #163+
      ENDIF.                                                                                                " issue #163+
*    lo_range = ip_excel->add_new_range( ).                                                               " issue #163-
      IF lo_range IS BOUND.                                                                                 "ins issue #235 - repeat rows/cols
        lo_range->name = ls_range-name.
        lo_range->set_range_value( lv_range_value ).
      ENDIF.                                                                                                "ins issue #235 - repeat rows/cols
      lo_node ?= lo_node->get_next( ).

    ENDWHILE.

  ENDMETHOD.