  METHOD encrypt_password.

    DATA lv_curr_offset            TYPE i.
    DATA lv_curr_char              TYPE c LENGTH 1.
    DATA lv_curr_hex               TYPE zexcel_pwd_hash.
    DATA lv_pwd_len                TYPE zexcel_pwd_hash.
    DATA lv_pwd_hash               TYPE zexcel_pwd_hash.

    CONSTANTS:
      lv_0x7fff TYPE zexcel_pwd_hash VALUE '7FFF',
      lv_0x0001 TYPE zexcel_pwd_hash VALUE '0001',
      lv_0xce4b TYPE zexcel_pwd_hash VALUE 'CE4B'.

    DATA lv_pwd            TYPE zexcel_aes_password.

    lv_pwd = i_pwd.

    lv_pwd_len = strlen( lv_pwd ).
    lv_curr_offset = lv_pwd_len - 1.

    WHILE lv_curr_offset GE 0.

      lv_curr_char = lv_pwd+lv_curr_offset(1).
      lv_curr_hex = char2hex( lv_curr_char ).

      lv_pwd_hash = (  shr14( lv_pwd_hash ) BIT-AND lv_0x0001 ) BIT-OR ( shl01( lv_pwd_hash ) BIT-AND lv_0x7fff ).

      lv_pwd_hash = lv_pwd_hash BIT-XOR lv_curr_hex.
      SUBTRACT 1 FROM lv_curr_offset.
    ENDWHILE.

    lv_pwd_hash = (  shr14( lv_pwd_hash ) BIT-AND lv_0x0001 ) BIT-OR ( shl01( lv_pwd_hash ) BIT-AND lv_0x7fff ).
    lv_pwd_hash = lv_pwd_hash BIT-XOR lv_0xce4b.
    lv_pwd_hash = lv_pwd_hash BIT-XOR lv_pwd_len.

    WRITE lv_pwd_hash TO r_encrypted_pwd.

  ENDMETHOD.