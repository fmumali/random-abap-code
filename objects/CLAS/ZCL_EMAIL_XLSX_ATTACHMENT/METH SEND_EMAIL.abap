  METHOD send_email.
  "Get data
  SELECT * FROM /dmo/flight into table @data(lt_data).
  Get REFERENCE OF lt_data into data(lr_data_ref).
  data(ls_xstring) = NEW zcl_itab_to_excel( )->itab_to_xstring( lr_data_Ref ).

  "Create email.
  try.
   "Create send request
        DATA(lr_send_request) = cl_bcs=>create_persistent( ).

        "Create mail body
        DATA(lt_body) = VALUE bcsy_text(
                          ( line = 'Dear Recipient,' ) ( )
                          ( line = 'PFA flight details file.' ) ( )
                          ( line = 'Thank You' )
                        ).

        "Set up document object
        DATA(lr_document) = cl_document_bcs=>create_document( i_type    = 'RAW'
                                                              i_text    = lt_body
                                                              i_subject = 'Flight Details' ).


        "Add attachment
        lr_document->add_attachment( i_attachment_type    = 'xls'
                                     i_attachment_size    = CONV #( xstrlen( ls_xstring ) )
                                     i_attachment_subject = 'Flight Details'
                                     i_attachment_header  = VALUE #( ( line = 'Flights.xlsx' ) )
                                     i_att_content_hex    = cl_bcs_convert=>xstring_to_solix( ls_xstring ) ).

        "Add document to send request
        lr_send_request->set_document( lr_document ).

        "Set sender
        lr_send_request->set_sender( cl_cam_address_bcs=>create_internet_address( i_address_string = CONV #( 'sender@dummy.com' ) ) ).


        "Set Recipient | This method has options to set CC/BCC as well
        lr_send_request->add_recipient(  i_recipient      = cl_cam_address_bcs=>create_internet_address(
                                         i_address_string = CONV #( 'recipient@dummy.com' )  )
          i_express   = abap_true ).


        "Send Email
        DATA(lv_sent_to_all) = lr_send_request->send( ).
        COMMIT WORK.


      CATCH cx_send_req_bcs INTO DATA(lx_req_bsc).
        "Error handling
      CATCH cx_document_bcs INTO DATA(lx_doc_bcs).
        "Error handling
      CATCH cx_address_bcs  INTO DATA(lx_add_bcs).
        "Error handling
    ENDTRY.
  ENDMETHOD.