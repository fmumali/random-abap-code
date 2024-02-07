  METHOD build_gradient.
    CHECK check_filltype_is_gradient( ) EQ abap_true.
    CLEAR gradtype.
    CASE filltype.
      WHEN c_fill_gradient_horizontal90.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_horizontal270.
        gradtype-degree = '270'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_horizontalb.
        gradtype-degree = '90'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_vertical.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_fromcenter.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
      WHEN c_fill_gradient_diagonal45.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_diagonal45b.
        gradtype-degree = '45'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_diagonal135.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_diagonal135b.
        gradtype-degree = '135'.
        gradtype-position1 = '0'.
        gradtype-position2 = '0.5'.
        gradtype-position3 = '1'.
      WHEN c_fill_gradient_cornerlt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
      WHEN c_fill_gradient_cornerlb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '1'.
        gradtype-top = '1'.
      WHEN c_fill_gradient_cornerrt.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-left = '1'.
        gradtype-right = '1'.
      WHEN c_fill_gradient_cornerrb.
        gradtype-type = c_fill_gradient_path.
        gradtype-position1 = '0'.
        gradtype-position2 = '1'.
        gradtype-bottom = '0.5'.
        gradtype-top = '0.5'.
        gradtype-left = '0.5'.
        gradtype-right = '0.5'.
    ENDCASE.

  ENDMETHOD.                    "build_gradient