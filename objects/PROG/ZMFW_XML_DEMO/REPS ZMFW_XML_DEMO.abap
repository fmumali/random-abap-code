*&---------------------------------------------------------------------*
*& Report ZMFW_XML_DEMO
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zmfw_xml_demo.
*Declarations to create XML document
DATA: lr_xml    TYPE REF TO if_ixml,
      file_path TYPE string   VALUE 'C:\Users\fredr\fmumali\projects\sap\zmfw\manifest.xml'.

*Create iXML object and document
DATA(lr_ixml) = cl_ixml=>create( ).
DATA(lr_document) = lr_ixml->create_document( ).

*Create and set encoding
lr_document->set_encoding( lr_ixml->create_encoding( byte_order = 0 character_set = 'utf-8' ) ).

*Create element "extension" as root
DATA(lr_extension) = lr_document->create_simple_element( name   = 'extension'
                                                         parent = lr_document  ).
"Set attribute
lr_extension->set_attribute( name  = 'Type'
                             value = 'Component' ).

*Create files element with extension element as parent
DATA(lr_files) = lr_document->create_simple_element( name   = 'files'
                                                    parent  = lr_extension ).
"Set attributes
lr_files->set_attribute( name  = 'folder'
                         value = 'site' ).

"Create element content
lr_document->create_simple_element( name   = 'filename'
                                    parent = lr_files
                                    value  = 'index.html' ).

lr_document->create_simple_element( name   = 'filename'
                                    parent = lr_files
                                    value  = 'site.php' ).

"Create media element
DATA(lr_media) = lr_document->create_simple_element( name   = 'media'
                                                     parent = lr_extension ).
lr_media->set_Attribute( name = 'folder'
                         value = 'media' ).

lr_document->create_simple_element( name   = 'folder'
                                    parent = lr_media
                                    value  = 'css' ).

lr_document->create_simple_element( name   = 'folder'
                                    parent = lr_media
                                    value  = 'images' ).

lr_document->create_simple_element( name   = 'folder'
                                    parent = lr_media
                                    value  = 'js' ).

"Create stream factory and render
DATA(lr_streamfactory) = lr_ixml->create_stream_factory( ).
DATA(lr_ostream) = lr_streamfactory->create_ostream_uri( system_id = file_path ).
DATA(lr_renderer) = lr_ixml->create_renderer( ostream = lr_ostream
                                            document = lr_document ).
*Set pretty print
lr_ostream->set_pretty_print( abap_true ).
lr_renderer->render( ).