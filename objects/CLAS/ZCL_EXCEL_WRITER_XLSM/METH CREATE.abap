  METHOD create.


* Office 2007 file format is a cab of several xml files with extension .xlsx

    DATA: lo_zip              TYPE REF TO cl_abap_zip,
          lo_worksheet        TYPE REF TO zcl_excel_worksheet,
          lo_active_worksheet TYPE REF TO zcl_excel_worksheet,
          lo_iterator         TYPE REF TO zcl_excel_collection_iterator,
          lo_nested_iterator  TYPE REF TO zcl_excel_collection_iterator,
          lo_table            TYPE REF TO zcl_excel_table,
          lo_drawing          TYPE REF TO zcl_excel_drawing,
          lo_drawings         TYPE REF TO zcl_excel_drawings.

    DATA: lv_content         TYPE xstring,
          lv_active          TYPE flag,
          lv_xl_sheet        TYPE string,
          lv_xl_sheet_rels   TYPE string,
          lv_xl_drawing      TYPE string,
          lv_xl_drawing_rels TYPE string,
          lv_syindex         TYPE string,
          lv_value           TYPE string,
          lv_drawing_index   TYPE i,
          lv_comment_index   TYPE i. " (+) Issue 588

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
    lv_drawing_index = 1.

    WHILE lo_iterator->has_next( ) EQ abap_true.
      lo_worksheet ?= lo_iterator->get_next( ).
      IF lo_active_worksheet->get_guid( ) EQ lo_worksheet->get_guid( ).
        lv_active = abap_true.
      ELSE.
        lv_active = abap_false.
      ENDIF.

      lv_content = me->create_xl_sheet( io_worksheet = lo_worksheet
                                        iv_active    = lv_active ).
      lv_xl_sheet = me->c_xl_sheet.
      lv_syindex = sy-index.
      lv_comment_index = sy-index. " (+) Issue 588
      SHIFT lv_syindex RIGHT DELETING TRAILING space.
      SHIFT lv_syindex LEFT DELETING LEADING space.
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet WITH lv_syindex.
      lo_zip->add( name    = lv_xl_sheet
                   content = lv_content ).

      lv_xl_sheet_rels = me->c_xl_sheet_rels.
      lv_content = me->create_xl_sheet_rels( io_worksheet = lo_worksheet
                                             iv_drawing_index = lv_drawing_index
                                             iv_comment_index = lv_comment_index ). " (+) Issue 588
      REPLACE ALL OCCURRENCES OF '#' IN lv_xl_sheet_rels WITH lv_syindex.
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

* Add drawings **********************************
      lo_drawings = lo_worksheet->get_drawings( ).
      IF lo_drawings->is_empty( ) = abap_false.
        lv_syindex = lv_drawing_index.
        SHIFT lv_syindex RIGHT DELETING TRAILING space.
        SHIFT lv_syindex LEFT DELETING LEADING space.

        lv_content = me->create_xl_drawings( lo_worksheet ).
        lv_xl_drawing = me->c_xl_drawings.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing WITH lv_syindex.
        lo_zip->add( name    = lv_xl_drawing
                     content = lv_content ).

        lv_content = me->create_xl_drawings_rels( lo_worksheet ).
        lv_xl_drawing_rels = me->c_xl_drawings_rels.
        REPLACE ALL OCCURRENCES OF '#' IN lv_xl_drawing_rels WITH lv_syindex.
        lo_zip->add( name    = lv_xl_drawing_rels
                     content = lv_content ).
        ADD 1 TO lv_drawing_index.
      ENDIF.
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
      lv_value = lo_drawing->get_media_name( ).
      CONCATENATE 'xl/charts/' lv_value INTO lv_value.
      lo_zip->add( name    = lv_value
                   content = lv_content ).
    ENDWHILE.

**********************************************************************
* STEP 9: Add vbaProject.bin to zip
    lo_zip->add( name    = me->c_xl_vbaproject
                 content = me->excel->zif_excel_book_vba_project~vbaproject ).

**********************************************************************
* STEP 12: Create the final zip
    ep_excel = lo_zip->save( ).

  ENDMETHOD.