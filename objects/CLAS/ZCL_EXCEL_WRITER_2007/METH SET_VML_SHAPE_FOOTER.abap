  METHOD set_vml_shape_footer.

    CONSTANTS: lc_shape               TYPE string VALUE '<v:shape id="{ID}" o:spid="_x0000_s1025" type="#_x0000_t75" style=''position:absolute;margin-left:0;margin-top:0;width:{WIDTH}pt;height:{HEIGHT}pt; z-index:1''>',
               lc_shape_image         TYPE string VALUE '<v:imagedata o:relid="{RID}" o:title="Logo Title"/><o:lock v:ext="edit" rotation="t"/></v:shape>',
               lc_shape_header_center TYPE string VALUE 'CH',
               lc_shape_header_left   TYPE string VALUE 'LH',
               lc_shape_header_right  TYPE string VALUE 'RH',
               lc_shape_footer_center TYPE string VALUE 'CF',
               lc_shape_footer_left   TYPE string VALUE 'LF',
               lc_shape_footer_right  TYPE string VALUE 'RF'.

    DATA: lv_content_left         TYPE string,
          lv_content_center       TYPE string,
          lv_content_right        TYPE string,
          lv_content_image_left   TYPE string,
          lv_content_image_center TYPE string,
          lv_content_image_right  TYPE string,
          lv_value                TYPE string,
          ls_drawing_position     TYPE zexcel_drawing_position.

    IF is_footer-left_image IS NOT INITIAL.
      lv_content_left = lc_shape.
      REPLACE '{ID}' IN lv_content_left WITH lc_shape_footer_left.
      ls_drawing_position = is_footer-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_left WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_left WITH lv_value.
      lv_content_image_left = lc_shape_image.
      lv_value = is_footer-left_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_left WITH lv_value.
    ENDIF.
    IF is_footer-center_image IS NOT INITIAL.
      lv_content_center = lc_shape.
      REPLACE '{ID}' IN lv_content_center WITH lc_shape_footer_center.
      ls_drawing_position = is_footer-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_center WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_center WITH lv_value.
      lv_content_image_center = lc_shape_image.
      lv_value = is_footer-center_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_center WITH lv_value.
    ENDIF.
    IF is_footer-right_image IS NOT INITIAL.
      lv_content_right = lc_shape.
      REPLACE '{ID}' IN lv_content_right WITH lc_shape_footer_right.
      ls_drawing_position = is_footer-left_image->get_position( ).
      IF ls_drawing_position-size-height IS NOT INITIAL.
        lv_value = ls_drawing_position-size-height.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{HEIGHT}' IN lv_content_right WITH lv_value.
      IF ls_drawing_position-size-width IS NOT INITIAL.
        lv_value = ls_drawing_position-size-width.
      ELSE.
        lv_value = '100'.
      ENDIF.
      CONDENSE lv_value.
      REPLACE '{WIDTH}' IN lv_content_right WITH lv_value.
      lv_content_image_right = lc_shape_image.
      lv_value = is_footer-right_image->get_index( ).
      CONDENSE lv_value.
      CONCATENATE 'rId' lv_value INTO lv_value.
      REPLACE '{RID}' IN lv_content_image_right WITH lv_value.
    ENDIF.

    CONCATENATE lv_content_left
                lv_content_image_left
                lv_content_center
                lv_content_image_center
                lv_content_right
                lv_content_image_right
           INTO ep_content.

  ENDMETHOD.