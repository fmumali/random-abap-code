  METHOD shift_formula.

    CONSTANTS: lcv_operators            TYPE string VALUE '+-/*^%=<>&, !',
               lcv_letters              TYPE string VALUE 'ABCDEFGHIJKLMNOPQRSTUVWXYZ$',
               lcv_digits               TYPE string VALUE '0123456789',
               lcv_cell_reference_error TYPE string VALUE '#REF!'.

    DATA: lv_tcnt          TYPE i,         " Counter variable
          lv_tlen          TYPE i,         " Temp variable length
          lv_cnt           TYPE i,         " Counter variable
          lv_cnt2          TYPE i,         " Counter variable
          lv_offset1       TYPE i,         " Character offset
          lv_numchars      TYPE i,         " Number of characters counter
          lv_tchar(1)      TYPE c,         " Temp character
          lv_tchar2(1)     TYPE c,         " Temp character
          lv_cur_form      TYPE string,    " Formula for current cell
          lv_ref_cell_addr TYPE string,    " Reference cell address
          lv_tcol1         TYPE string,    " Temp column letter
          lv_tcol2         TYPE string,    " Temp column letter
          lv_tcoln         TYPE i,         " Temp column number
          lv_trow1         TYPE string,    " Temp row number
          lv_trow2         TYPE string,    " Temp row number
          lv_flen          TYPE i,         " Length of reference formula
          lv_tlen2         TYPE i,         " Temp variable length
          lv_substr1       TYPE string,    " Substring variable
          lv_abscol        TYPE string,    " Absolute column symbol
          lv_absrow        TYPE string,    " Absolute row symbol
          lv_ref_formula   TYPE string,
          lv_compare_1     TYPE string,
          lv_compare_2     TYPE string,
          lv_level         TYPE i,         " Level of groups [..[..]..] or {..}

          lv_errormessage  TYPE string.

*--------------------------------------------------------------------*
* When copying a cell in EXCEL to another cell any inherent formulas
* are copied as well.  Cell-references in the formula are being adjusted
* by the distance of the new cell to the original one
*--------------------------------------------------------------------*
* §1 Parse reference formula character by character
* §2 Identify Cell-references
* §3 Shift cell-reference
* §4 Build resulting formula
*--------------------------------------------------------------------*

    lv_ref_formula = iv_reference_formula.
*--------------------------------------------------------------------*
* No distance --> Reference = resulting cell/formula
*--------------------------------------------------------------------*
    IF    iv_shift_cols = 0
      AND iv_shift_rows = 0.
      ev_resulting_formula = lv_ref_formula.
      RETURN. " done
    ENDIF.


    lv_flen     = strlen( lv_ref_formula ).
    lv_numchars = 1.

*--------------------------------------------------------------------*
* §1 Parse reference formula character by character
*--------------------------------------------------------------------*
    DO lv_flen TIMES.

      CLEAR: lv_tchar,
             lv_substr1,
             lv_ref_cell_addr.
      lv_cnt2 = lv_cnt + 1.
      IF lv_cnt2 > lv_flen.
        EXIT. " Done
      ENDIF.

*--------------------------------------------------------------------*
* Here we have the current character in the formula
*--------------------------------------------------------------------*
      lv_tchar = lv_ref_formula+lv_cnt(1).

*--------------------------------------------------------------------*
* Operators or opening parenthesis will separate possible cellreferences
*--------------------------------------------------------------------*
      IF    (    lv_tchar CA lcv_operators
              OR lv_tchar CA '(' )
        AND lv_cnt2 = 1.
        lv_substr1  = lv_ref_formula+lv_offset1(1).
        CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
        lv_cnt      = lv_cnt + 1.
        lv_offset1  = lv_cnt.
        lv_numchars = 1.
        CONTINUE.       " --> next character in formula can be analyzed
      ENDIF.

*--------------------------------------------------------------------*
* Quoted literal text holds no cell reference --> advance to end of text
*--------------------------------------------------------------------*
      IF lv_tchar EQ '"'.
        lv_cnt      = lv_cnt + 1.
        lv_numchars = lv_numchars + 1.
        lv_tchar     = lv_ref_formula+lv_cnt(1).
        WHILE lv_tchar NE '"'.

          lv_cnt      = lv_cnt + 1.
          lv_numchars = lv_numchars + 1.
          lv_tchar    = lv_ref_formula+lv_cnt(1).

        ENDWHILE.
        lv_cnt2    = lv_cnt + 1.
        lv_substr1 = lv_ref_formula+lv_offset1(lv_numchars).
        CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
        lv_cnt     = lv_cnt + 1.
        IF lv_cnt = lv_flen.
          EXIT.
        ENDIF.
        lv_offset1  = lv_cnt.
        lv_numchars = 1.
        lv_tchar    = lv_ref_formula+lv_cnt(1).
        lv_cnt2     = lv_cnt + 1.
        CONTINUE.       " --> next character in formula can be analyzed
      ENDIF.


*--------------------------------------------------------------------*
* Groups - Ignore values inside blocks [..[..]..] and {..}
*     R1C1-Style Cell Reference: R[1]C[1]
*     Cell References: 'C:\[Source.xlsx]Sheet1'!$A$1
*     Array constants: {1,3.5,TRUE,"Hello"}
*     "Intra table reference": Flights[[#This Row],[Air fare]]
*--------------------------------------------------------------------*
      IF lv_tchar CA '[]{}' OR lv_level > 0.
        IF lv_tchar CA '[{'.
          lv_level = lv_level + 1.
        ELSEIF lv_tchar CA ']}'.
          lv_level = lv_level - 1.
        ENDIF.
        IF lv_cnt2 = lv_flen.
          lv_substr1 = iv_reference_formula+lv_offset1(lv_numchars).
          CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
          EXIT.
        ENDIF.
        lv_numchars = lv_numchars + 1.
        lv_cnt   = lv_cnt   + 1.
        lv_cnt2  = lv_cnt   + 1.
        CONTINUE.
      ENDIF.

*--------------------------------------------------------------------*
* Operators or parenthesis or last character in formula will separate possible cellreferences
*--------------------------------------------------------------------*
      IF   lv_tchar CA lcv_operators
        OR lv_tchar CA '():'
        OR lv_cnt2  =  lv_flen.
        IF lv_cnt > 0.
          lv_substr1 = lv_ref_formula+lv_offset1(lv_numchars).
*--------------------------------------------------------------------*
* Check for text concatenation and functions
*--------------------------------------------------------------------*
          IF ( lv_tchar CA lcv_operators AND lv_tchar EQ lv_substr1 ) OR lv_tchar EQ '('.
            CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            CONTINUE.       " --> next character in formula can be analyzed
          ENDIF.

          lv_tlen = lv_cnt2 - lv_offset1.
*--------------------------------------------------------------------*
* Exclude mathematical operators and closing parentheses
*--------------------------------------------------------------------*
          IF   lv_tchar CA lcv_operators
            OR lv_tchar CA ':)'.
            IF    lv_cnt2     = lv_flen
              AND lv_numchars = 1.
              CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
              lv_cnt      = lv_cnt + 1.
              lv_offset1  = lv_cnt.
              lv_cnt2     = lv_cnt + 1.
              lv_numchars = 1.
              CONTINUE.       " --> next character in formula can be analyzed
            ELSE.
              lv_tlen = lv_tlen - 1.
            ENDIF.
          ENDIF.
*--------------------------------------------------------------------*
* Capture reference cell address
*--------------------------------------------------------------------*
          TRY.
              lv_ref_cell_addr = lv_ref_formula+lv_offset1(lv_tlen). "Ref cell address
            CATCH cx_root.
              lv_errormessage = 'Internal error in Class ZCL_EXCEL_COMMON Method SHIFT_FORMULA Spot 1 '.  " Change to messageclass if possible
              zcx_excel=>raise_text( lv_errormessage ).
          ENDTRY.

*--------------------------------------------------------------------*
* Split cell address into characters and numbers
*--------------------------------------------------------------------*
          CLEAR: lv_tlen,
                 lv_tcnt,
                 lv_tcol1,
                 lv_trow1.
          lv_tlen = strlen( lv_ref_cell_addr ).
          IF lv_tlen <> 0.
            CLEAR: lv_tcnt.
            DO lv_tlen TIMES.
              CLEAR: lv_tchar2.
              lv_tchar2 = lv_ref_cell_addr+lv_tcnt(1).
              IF lv_tchar2 CA lcv_letters.
                CONCATENATE lv_tcol1 lv_tchar2 INTO lv_tcol1.
              ELSEIF lv_tchar2 CA lcv_digits.
                CONCATENATE lv_trow1 lv_tchar2 INTO lv_trow1.
              ENDIF.
              lv_tcnt = lv_tcnt + 1.
            ENDDO.
          ENDIF.

          " Is valid column & row ?
          IF lv_tcol1 IS NOT INITIAL AND lv_trow1 IS NOT INITIAL.
            " COLUMN + ROW
            CONCATENATE lv_tcol1 lv_trow1 INTO lv_compare_1.
            " Original condensed string
            lv_compare_2 = lv_ref_cell_addr.
            CONDENSE lv_compare_2.
            IF lv_compare_1 <> lv_compare_2.
              CLEAR: lv_trow1, lv_tchar2.
            ENDIF.
          ENDIF.

*--------------------------------------------------------------------*
* Check for invalid cell address
*--------------------------------------------------------------------*
          IF lv_tcol1 IS INITIAL OR lv_trow1 IS INITIAL.
            CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            CONTINUE.
          ENDIF.
*--------------------------------------------------------------------*
* Check for range names
*--------------------------------------------------------------------*
          CLEAR: lv_tlen.
          lv_tlen = strlen( lv_tcol1 ).
          IF lv_tlen GT 3.
            CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            CONTINUE.
          ENDIF.
*--------------------------------------------------------------------*
* Check for valid row
*--------------------------------------------------------------------*
          IF lv_trow1 GT 1048576.
            CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
            lv_cnt = lv_cnt + 1.
            lv_offset1 = lv_cnt.
            lv_cnt2 = lv_cnt + 1.
            lv_numchars = 1.
            CONTINUE.
          ENDIF.
*--------------------------------------------------------------------*
* Check for absolute column or row reference
*--------------------------------------------------------------------*
          CLEAR: lv_tcol2,
                 lv_trow2,
                 lv_abscol,
                 lv_absrow.
          lv_tlen2 = strlen( lv_tcol1 ) - 1.
          IF lv_tcol1 IS NOT INITIAL.
            lv_abscol = lv_tcol1(1).
          ENDIF.
          IF lv_tlen2 GE 0.
            lv_absrow = lv_tcol1+lv_tlen2(1).
          ENDIF.
          IF lv_abscol EQ '$' AND lv_absrow EQ '$'.
            lv_tlen2 = lv_tlen2 - 1.
            IF lv_tlen2 > 0.
              lv_tcol1 = lv_tcol1+1(lv_tlen2).
            ENDIF.
            lv_tlen2 = lv_tlen2 + 1.
          ELSEIF lv_abscol EQ '$'.
            lv_tcol1 = lv_tcol1+1(lv_tlen2).
          ELSEIF lv_absrow EQ '$'.
            lv_tcol1 = lv_tcol1(lv_tlen2).
          ENDIF.
*--------------------------------------------------------------------*
* Check for valid column
*--------------------------------------------------------------------*
          TRY.
              lv_tcoln = zcl_excel_common=>convert_column2int( lv_tcol1 ) + iv_shift_cols.
            CATCH zcx_excel.
              CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
              lv_cnt = lv_cnt + 1.
              lv_offset1 = lv_cnt.
              lv_cnt2 = lv_cnt + 1.
              lv_numchars = 1.
              CONTINUE.
          ENDTRY.
*--------------------------------------------------------------------*
* Check whether there is a referencing problem
*--------------------------------------------------------------------*
          lv_trow2 = lv_trow1 + iv_shift_rows.
          " Remove the space used for the sign
          CONDENSE lv_trow2.
          IF   ( lv_tcoln < 1 AND lv_abscol <> '$' )   " Maybe we should add here max-column and max row-tests as well.
            OR ( lv_trow2 < 1 AND lv_absrow <> '$' ).  " Check how EXCEL behaves in this case
*--------------------------------------------------------------------*
* Referencing problem encountered --> set error
*--------------------------------------------------------------------*
            CONCATENATE lv_cur_form lcv_cell_reference_error INTO lv_cur_form.
          ELSE.
*--------------------------------------------------------------------*
* No referencing problems --> adjust row and column
*--------------------------------------------------------------------*

*--------------------------------------------------------------------*
* Adjust column
*--------------------------------------------------------------------*
            IF lv_abscol EQ '$'.
              CONCATENATE lv_cur_form lv_abscol lv_tcol1 INTO lv_cur_form.
            ELSEIF iv_shift_cols EQ 0.
              CONCATENATE lv_cur_form lv_tcol1 INTO lv_cur_form.
            ELSE.
              TRY.
                  lv_tcol2 = zcl_excel_common=>convert_column2alpha( lv_tcoln ).
                  CONCATENATE lv_cur_form lv_tcol2 INTO lv_cur_form.
                CATCH zcx_excel.
                  CONCATENATE lv_cur_form lv_substr1 INTO lv_cur_form.
                  lv_cnt = lv_cnt + 1.
                  lv_offset1 = lv_cnt.
                  lv_cnt2 = lv_cnt + 1.
                  lv_numchars = 1.
                  CONTINUE.
              ENDTRY.
            ENDIF.
*--------------------------------------------------------------------*
* Adjust row
*--------------------------------------------------------------------*
            IF lv_absrow EQ '$'.
              CONCATENATE lv_cur_form lv_absrow lv_trow1 INTO lv_cur_form.
            ELSEIF iv_shift_rows = 0.
              CONCATENATE lv_cur_form lv_trow1 INTO lv_cur_form.
            ELSE.
              CONCATENATE lv_cur_form lv_trow2 INTO lv_cur_form.
            ENDIF.
          ENDIF.

          lv_numchars = 0.
          IF   lv_tchar CA lcv_operators
            OR lv_tchar CA ':)'.
            CONCATENATE lv_cur_form lv_tchar INTO lv_cur_form RESPECTING BLANKS.
          ENDIF.
          lv_offset1 = lv_cnt2.
        ENDIF.
      ENDIF.
      lv_numchars = lv_numchars + 1.
      lv_cnt   = lv_cnt   + 1.
      lv_cnt2  = lv_cnt   + 1.

    ENDDO.



*--------------------------------------------------------------------*
* Return resulting formula
*--------------------------------------------------------------------*
    IF lv_cur_form IS NOT INITIAL.
      ev_resulting_formula = lv_cur_form.
    ENDIF.

  ENDMETHOD.