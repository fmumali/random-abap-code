  METHOD fill_struct_from_attributes.
*--------------------------------------------------------------------*
* issue #230   - Pimp my Code
*              - Stefan Schmoecker,      (done)              2012-11-07
*              - ...
* changes: renaming variables to naming conventions
*          aligning code
*          adding comments to explain what we are trying to achieve
*--------------------------------------------------------------------*

    DATA: lv_name       TYPE string,
          lo_attributes TYPE REF TO if_ixml_named_node_map,
          lo_attribute  TYPE REF TO if_ixml_attribute,
          lo_iterator   TYPE REF TO if_ixml_node_iterator.

    FIELD-SYMBOLS: <component>                  TYPE any.

*--------------------------------------------------------------------*
* The values of named attributes of a tag are being read and moved into corresponding
* fields of given structure
* Behaves like move-corresonding tag to structure

* Example:
*     <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>
*   Here the attributes are Target, Type and Id.  Thus if the passed
*   structure has fieldnames Id and Target these would be filled with
*   "rId3" and "docProps/app.xml" respectively
*--------------------------------------------------------------------*
    CLEAR cp_structure.

    lo_attributes  = ip_element->get_attributes( ).
    lo_iterator    = lo_attributes->create_iterator( ).
    lo_attribute  ?= lo_iterator->get_next( ).
    WHILE lo_attribute IS BOUND.

      lv_name = lo_attribute->get_name( ).
      TRANSLATE lv_name TO UPPER CASE.
      ASSIGN COMPONENT lv_name OF STRUCTURE cp_structure TO <component>.
      IF sy-subrc = 0.
        <component> = lo_attribute->get_value( ).
      ENDIF.
      lo_attribute ?= lo_iterator->get_next( ).

    ENDWHILE.


  ENDMETHOD.