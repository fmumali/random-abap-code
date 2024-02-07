  METHOD shr14.

    DATA:
      lv_bit      TYPE i,
      lv_curr_pos TYPE i,
      lv_next_pos TYPE i.

    r_pwd_hash = i_pwd_hash.

    DO 14 TIMES.
      lv_curr_pos = 15.
      lv_next_pos = 16.

      DO 15 TIMES.
        GET BIT lv_curr_pos OF r_pwd_hash INTO lv_bit.
        SET BIT lv_next_pos OF r_pwd_hash TO lv_bit.
        SUBTRACT 1 FROM lv_curr_pos.
        SUBTRACT 1 FROM lv_next_pos.
      ENDDO.
      SET BIT 1 OF r_pwd_hash TO 0.
    ENDDO.

  ENDMETHOD.