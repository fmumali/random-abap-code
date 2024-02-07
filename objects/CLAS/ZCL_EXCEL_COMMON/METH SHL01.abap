  METHOD shl01.

    DATA:
      lv_bit      TYPE i,
      lv_curr_pos TYPE i VALUE 2,
      lv_prev_pos TYPE i VALUE 1.

    DO 15 TIMES.
      GET BIT lv_curr_pos OF i_pwd_hash INTO lv_bit.
      SET BIT lv_prev_pos OF r_pwd_hash TO lv_bit.
      ADD 1 TO lv_curr_pos.
      ADD 1 TO lv_prev_pos.
    ENDDO.
    SET BIT 16 OF r_pwd_hash TO 0.

  ENDMETHOD.