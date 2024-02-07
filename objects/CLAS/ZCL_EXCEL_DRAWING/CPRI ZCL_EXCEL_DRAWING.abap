  PRIVATE SECTION.

*"* private components of class ZCL_EXCEL_DRAWING
*"* do not include other source files here!!!
    DATA type TYPE zexcel_drawing_type VALUE type_image. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    DATA index TYPE string .
    DATA anchor TYPE zexcel_drawing_anchor VALUE anchor_one_cell. "#EC NOTEXT .  .  .  .  .  .  .  .  .  .  . " .
    CONSTANTS c_media_source_www TYPE c VALUE 1.            "#EC NOTEXT
    CONSTANTS c_media_source_xstring TYPE c VALUE 0.        "#EC NOTEXT
    CONSTANTS c_media_source_mime TYPE c VALUE 2.           "#EC NOTEXT
    DATA guid TYPE zexcel_guid .
    DATA media TYPE xstring .
    DATA media_key_www TYPE wwwdatatab .
    DATA media_name TYPE string .
    DATA media_source TYPE c .
    DATA media_type TYPE string .
    DATA io TYPE skwf_io .
    DATA from_loc TYPE zexcel_drawing_location .
    DATA to_loc TYPE zexcel_drawing_location .
    DATA size TYPE zexcel_drawing_size .
    CONSTANTS c_ixml_iid_element TYPE i VALUE 130.
