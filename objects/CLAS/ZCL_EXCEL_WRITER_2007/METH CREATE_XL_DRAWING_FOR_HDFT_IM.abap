  METHOD create_xl_drawing_for_hdft_im.


    DATA:
      ld_1            TYPE string,
      ld_2            TYPE string,
      ld_3            TYPE string,
      ld_4            TYPE string,
      ld_5            TYPE string,
      ld_7            TYPE string,

      ls_odd_header   TYPE zexcel_s_worksheet_head_foot,
      ls_odd_footer   TYPE zexcel_s_worksheet_head_foot,
      ls_even_header  TYPE zexcel_s_worksheet_head_foot,
      ls_even_footer  TYPE zexcel_s_worksheet_head_foot,
      lv_content      TYPE string.


* INIT_RESULT
    CLEAR ep_content.


* BODY
    ld_1 = '<xml xmlns:v="urn:schemas-microsoft-com:vml"  xmlns:o="urn:schemas-microsoft-com:office:office"  xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'.
    ld_2 = '<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/>'.
    ld_3 = '<v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/>'.
    ld_4 = '<v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>'.


    CONCATENATE ld_1
                ld_2
                ld_3
                ld_4
         INTO lv_content.

    io_worksheet->sheet_setup->get_header_footer( IMPORTING ep_odd_header = ls_odd_header
                                                            ep_odd_footer = ls_odd_footer
                                                            ep_even_header = ls_even_header
                                                            ep_even_footer = ls_even_footer ).

    ld_5 = me->set_vml_shape_header( ls_odd_header ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.
    ld_5 = me->set_vml_shape_header( ls_even_header ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.
    ld_5 = me->set_vml_shape_footer( ls_odd_footer ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.
    ld_5 = me->set_vml_shape_footer( ls_even_footer ).
    CONCATENATE lv_content
                ld_5
           INTO lv_content.

    ld_7 = '</xml>'.

    CONCATENATE lv_content
                ld_7
           INTO lv_content.

    CALL FUNCTION 'SCMS_STRING_TO_XSTRING'
      EXPORTING
        text   = lv_content
      IMPORTING
        buffer = ep_content
      EXCEPTIONS
        failed = 1
        OTHERS = 2.
    IF sy-subrc <> 0.
      CLEAR ep_content.
    ENDIF.

  ENDMETHOD.