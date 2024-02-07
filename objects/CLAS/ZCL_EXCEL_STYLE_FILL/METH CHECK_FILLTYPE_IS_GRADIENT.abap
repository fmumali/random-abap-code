  METHOD check_filltype_is_gradient.
    CASE filltype.
      WHEN c_fill_gradient_horizontal90 OR
           c_fill_gradient_horizontal270 OR
           c_fill_gradient_horizontalb OR
           c_fill_gradient_vertical OR
           c_fill_gradient_fromcenter OR
           c_fill_gradient_diagonal45 OR
           c_fill_gradient_diagonal45b OR
           c_fill_gradient_diagonal135 OR
           c_fill_gradient_diagonal135b OR
           c_fill_gradient_cornerlt OR
           c_fill_gradient_cornerlb OR
           c_fill_gradient_cornerrt OR
           c_fill_gradient_cornerrb.
        rv_is_gradient = abap_true.
    ENDCASE.
  ENDMETHOD.                    "check_filltype_is_gradient