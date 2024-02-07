  METHOD create.

* Office 2007 file format is a cab of several xml files with extension .xlsx

    DATA: lo_zip              TYPE REF TO cl_abap_zip,
          lo_worksheet        TYPE REF TO zcl_excel_worksheet,
          lo_active_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lo_nested_iterator  TYPE REF TO zcl_excel_collection_iterator,
          lo_table            TYPE REF TO zcl_excel_table,
          lo_drawing          TYPE REF TO zcl_excel_drawing,
          lo_drawings         TYPE REF TO zcl_excel_drawings,
          lo_comment          TYPE REF TO zcl_excel_comment,   " (+) Issue #180
          lo_comments         TYPE REF TO zcl_excel_comments.  " (+) Issue #180

    DATA: lv_content                TYPE xstring,
          lv_active                 TYPE flag,
          lv_xl_sheet               TYPE string,
          lv_xl_sheet_rels          TYPE string,
          lv_xl_drawing_for_comment TYPE string,   " (+) Issue #180
          lv_xl_comment             TYPE string,   " (+) Issue #180
          lv_xl_drawing             TYPE string,
          lv_xl_drawing_rels        TYPE string,
          lv_index_str              TYPE string,
          lv_value                  TYPE string,
          lv_sheet_index            TYPE i,
          lv_drawing_index          TYPE i,
          lv_comment_index          TYPE i.        " (+) Issue #180

**********************************************************************

**********************************************************************
* Start of insertion # issue 139 - Dateretention of cellstyles
    me->excel->add_static_styles( ).
* End of insertion # issue 139 - Dateretention of cellstyles

**********************************************************************
* STEP 1: Create archive object file (ZIP)
    CREATE OBJECT lo_zip.

**********************************************************************
* STEP 2: Add [Content_Types].xml to zip
    lv_content = me->create_content_types( ).
    lo_zip->add( name    = me->c_content_types
                 content = lv_content ).

**********************************************************************
* STEP 3: Add _rels/.rels to zip
    lv_content = me->create_relationships( ).
    lo_zip->add( name    = me->c_relationships
                 content = lv_content ).

**********************************************************************
* STEP 4: Add docProps/app.xml to zip
    lv_content = me->create_docprops_app( ).
    lo_zip->add( name    = me->c_docprops_app
                 content = lv_content ).

**********************************************************************
* STEP 5: Add docProps/core.xml to zip
    lv_content = me->create_docprops_core( ).
    lo_zip->add( name    = me->c_docprops_core
                 content = lv_content ).

**********************************************************************
* STEP 6: Add xl/_rels/workbook.xml.rels to zip
    lv_content = me->create_xl_relationships( ).
    lo_zip->add( name    = me->c_xl_relationships
                 content = lv_content ).

**********************************************************************
* STEP 6: Add xl/_rels/workbook.xml.rels to zip
    lv_content = me->create_xl_theme( ).
    lo_zip->add( name    = me->c_xl_theme
                 content = lv_content ).

**********************************************************************
* STEP 7: Add xl/workbook.xml to zip
    lv_content = me->create_xl_workbook( ).
    lo_zip->add( name    = me->c_xl_workbook
                 content = lv_content ).

**********************************************************************
* STEP 8: Add xl/workbook.xml to zip
    lv_content = me->create_xl_styles( ).
    lo_zip->add( name    = me->c_xl_styles
                 content = lv_content ).

**********************************************************************
* STEP 9: Add sharedStrings.xml to zip
    lv_content = me->create_xl_sharedstrings( ).
    lo_zip->add( name    = me->c_xl_sharedstrings
                 content = lv_content ).

**********************************************************************
* STEP 10: Add sheet#.xml and drawing#.xml to zip
    lo_iterator = me->excel->get_worksheets_iterator( ).
    lo_active_worksheet = me->excel->get_active_worksheet( ).

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lv_sheet_index = sy-index.

      lo_worksheet ?= lo_iterator->get_next( ).
      IF lo_active_worksheet->get_guid( ) EQ lo_worksheet->get_guid( ).
        lv_active = abap_true.
      ELSE.
        lv_active = abap_false.
      ENDIF.
      lv_content = me->create_xl_sheet( io_worksheet = lo_worksheet
                                        iv_active    = lv_active ).
      lv_xl_sheet = me->c_xl_sheet.

      lv_index_str = lv_sheet_index.
      CONDENSE lv_index_str NO-GAPS.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet WITH lv_index_str.
      lo_zip->add( name    = lv_xl_sheet
                   content = lv_content ).

* Begin - Add - Issue #180
* Add comments **********************************
      lo_comments = lo_worksheet->get_comments( ).
      IF lo_comments->is_empty( ) = abap_false.
        lv_comment_index = lv_comment_index + 1.

        " Create comment itself
        lv_content = me->create_xl_comments( lo_worksheet ).
        lv_xl_comment = me->c_xl_comments.
        lv_index_str = lv_comment_index.
        CONDENSE lv_index_str NO-GAPS.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_comment WITH lv_index_str.
        lo_zip->add( name    = lv_xl_comment
                     content = lv_content ).

        " Create vmlDrawing that will host the comment
        lv_content = me->create_xl_drawing_for_comments( lo_worksheet ).
        lv_xl_drawing_for_comment = me->cl_xl_drawing_for_comments.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_for_comment WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_for_comment
                     content = lv_content ).
      ENDIF.
* End   - Add - Issue #180

* Add drawings **********************************
      lo_drawings = lo_worksheet->get_drawings( ).
      IF lo_drawings->is_empty( ) = abap_false.
        lv_drawing_index = lv_drawing_index + 1.

        lv_content = me->create_xl_drawings( lo_worksheet ).
        lv_xl_drawing = me->c_xl_drawings.
        lv_index_str = lv_drawing_index.
        CONDENSE lv_index_str NO-GAPS.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing
                     content = lv_content ).

        lv_content = me->create_xl_drawings_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_rels.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_rels WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
      ENDIF.

* Add Header/Footer image
      DATA: lt_drawings TYPE zexcel_t_drawings.
      lt_drawings = lo_worksheet->get_header_footer_drawings( ).
      IF lines( lt_drawings ) > 0. "Header or footer image exist

        lv_comment_index = lv_comment_index + 1.
        lv_index_str = lv_comment_index.
        CONDENSE lv_index_str NO-GAPS.

        " Create vmlDrawing that will host the image
        lv_content = me->create_xl_drawing_for_hdft_im( lo_worksheet ).
        lv_xl_drawing_for_comment = me->cl_xl_drawing_for_comments.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_for_comment WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_for_comment
                     content = lv_content ).

        " Create vmlDrawing REL that will host the image
        lv_content = me->create_xl_drawings_hdft_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_vml_rels.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_rels WITH lv_index_str.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
      ENDIF.


      lv_xl_sheet_rels = me->c_xl_sheet_rels.
      lv_content = me->create_xl_sheet_rels( io_worksheet = lo_worksheet
                                             iv_drawing_index = lv_drawing_index
                                             iv_comment_index = lv_comment_index ).      " (+) Issue #180

      lv_index_str = lv_sheet_index.
      CONDENSE lv_index_str NO-GAPS.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet_rels WITH lv_index_str.
      lo_zip->add( name    = lv_xl_sheet_rels
                   content = lv_content ).

      lo_nested_iterator = lo_worksheet->get_tables_iterator( ).

      WHILE lo_nested_iterator->has_next( ) EQ abap_true.
        lo_table ?= lo_nested_iterator->get_next( ).
        lv_content = me->create_xl_table( lo_table ).

        lv_value = lo_table->get_name( ).
        CONCATENATE 'xl/tables/' lv_value '.xml' INTO lv_value.
        lo_zip->add( name = lv_value
                      content = lv_content ).
      ENDWHILE.



    ENDWHILE.

**********************************************************************
* STEP 11: Add media
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_image ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_content = lo_drawing->get_media( ).
      lv_value = lo_drawing->get_media_name( ).
      CONCATENATE 'xl/media/' lv_value INTO lv_value.
      lo_zip->add( name    = lv_value
                   content = lv_content ).
    ENDWHILE.

**********************************************************************
* STEP 12: Add charts
    lo_iterator = me->excel->get_drawings_iterator( zcl_excel_drawing=>type_chart ).
    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_drawing ?= lo_iterator->get_next( ).

      lv_content = lo_drawing->get_media( ).

      "-------------Added by Alessandro Iannacci - Only if template exist
      IF lv_content IS NOT INITIAL AND me->excel->use_template EQ abap_true.
        lv_value = lo_drawing->get_media_name( ).
        CONCATENATE 'xl/charts/' lv_value INTO lv_value.
        lo_zip->add( name    = lv_value
                     content = lv_content ).
      ELSE. "ADD CUSTOM CHART!!!!
        lv_content = me->create_xl_charts( lo_drawing ).
        lv_value = lo_drawing->get_media_name( ).
        CONCATENATE 'xl/charts/' lv_value INTO lv_value.
        lo_zip->add( name    = lv_value
                     content = lv_content ).
      ENDIF.
      "-------------------------------------------------
    ENDWHILE.

* Second to last step: Allow further information put into the zip archive by child classes
    me->add_further_data_to_zip( lo_zip ).

**********************************************************************
* Last step: Create the final zip
    ep_excel = lo_zip->save( ).

  ENDMETHOD.