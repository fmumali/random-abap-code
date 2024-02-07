CLASS zcl_excel_sheet_setup DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
*"* public components of class ZCL_EXCEL_SHEET_SETUP
*"* do not include other source files here!!!

    DATA black_and_white TYPE flag .
    DATA cell_comments TYPE string .
    DATA copies TYPE int2 .
    CONSTANTS c_break_column TYPE zexcel_break VALUE 2.     "#EC NOTEXT
    CONSTANTS c_break_none TYPE zexcel_break VALUE 0.       "#EC NOTEXT
    CONSTANTS c_break_row TYPE zexcel_break VALUE 1.        "#EC NOTEXT
    CONSTANTS c_cc_as_displayed TYPE string VALUE 'asDisplayed'. "#EC NOTEXT
    CONSTANTS c_cc_at_end TYPE string VALUE 'atEnd'.        "#EC NOTEXT
    CONSTANTS c_cc_none TYPE string VALUE 'none'.           "#EC NOTEXT
    CONSTANTS c_ord_downthenover TYPE string VALUE 'downThenOver'. "#EC NOTEXT
    CONSTANTS c_ord_overthendown TYPE string VALUE 'overThenDown'. "#EC NOTEXT
    CONSTANTS c_orientation_default TYPE zexcel_sheet_orienatation VALUE 'default'. "#EC NOTEXT
    CONSTANTS c_orientation_landscape TYPE zexcel_sheet_orienatation VALUE 'landscape'. "#EC NOTEXT
    CONSTANTS c_orientation_portrait TYPE zexcel_sheet_orienatation VALUE 'portrait'. "#EC NOTEXT
    CONSTANTS c_papersize_6_3_4_envelope TYPE zexcel_sheet_paper_size VALUE 38. "#EC NOTEXT
    CONSTANTS c_papersize_a2_paper TYPE zexcel_sheet_paper_size VALUE 64. "#EC NOTEXT
    CONSTANTS c_papersize_a3 TYPE zexcel_sheet_paper_size VALUE 8. "#EC NOTEXT
    CONSTANTS c_papersize_a3_extra_paper TYPE zexcel_sheet_paper_size VALUE 61. "#EC NOTEXT
    CONSTANTS c_papersize_a3_extra_tv_paper TYPE zexcel_sheet_paper_size VALUE 66. "#EC NOTEXT
    CONSTANTS c_papersize_a3_tv_paper TYPE zexcel_sheet_paper_size VALUE 65. "#EC NOTEXT
    CONSTANTS c_papersize_a4 TYPE zexcel_sheet_paper_size VALUE 9. "#EC NOTEXT
    CONSTANTS c_papersize_a4_extra_paper TYPE zexcel_sheet_paper_size VALUE 51. "#EC NOTEXT
    CONSTANTS c_papersize_a4_plus_paper TYPE zexcel_sheet_paper_size VALUE 58. "#EC NOTEXT
    CONSTANTS c_papersize_a4_small TYPE zexcel_sheet_paper_size VALUE 10. "#EC NOTEXT
    CONSTANTS c_papersize_a4_tv_paper TYPE zexcel_sheet_paper_size VALUE 53. "#EC NOTEXT
    CONSTANTS c_papersize_a5 TYPE zexcel_sheet_paper_size VALUE 11. "#EC NOTEXT
    CONSTANTS c_papersize_a5_extra_paper TYPE zexcel_sheet_paper_size VALUE 62. "#EC NOTEXT
    CONSTANTS c_papersize_a5_tv_paper TYPE zexcel_sheet_paper_size VALUE 59. "#EC NOTEXT
    CONSTANTS c_papersize_b4 TYPE zexcel_sheet_paper_size VALUE 12. "#EC NOTEXT
    CONSTANTS c_papersize_b4_envelope TYPE zexcel_sheet_paper_size VALUE 33. "#EC NOTEXT
    CONSTANTS c_papersize_b5 TYPE zexcel_sheet_paper_size VALUE 13. "#EC NOTEXT
    CONSTANTS c_papersize_b5_envelope TYPE zexcel_sheet_paper_size VALUE 34. "#EC NOTEXT
    CONSTANTS c_papersize_b6_envelope TYPE zexcel_sheet_paper_size VALUE 35. "#EC NOTEXT
    CONSTANTS c_papersize_c TYPE zexcel_sheet_paper_size VALUE 24. "#EC NOTEXT
    CONSTANTS c_papersize_c3_envelope TYPE zexcel_sheet_paper_size VALUE 29. "#EC NOTEXT
    CONSTANTS c_papersize_c4_envelope TYPE zexcel_sheet_paper_size VALUE 30. "#EC NOTEXT
    CONSTANTS c_papersize_c5_envelope TYPE zexcel_sheet_paper_size VALUE 28. "#EC NOTEXT
    CONSTANTS c_papersize_c65_envelope TYPE zexcel_sheet_paper_size VALUE 32. "#EC NOTEXT
    CONSTANTS c_papersize_c6_envelope TYPE zexcel_sheet_paper_size VALUE 31. "#EC NOTEXT
    CONSTANTS c_papersize_d TYPE zexcel_sheet_paper_size VALUE 25. "#EC NOTEXT
    CONSTANTS c_papersize_de_leg_fanfold TYPE zexcel_sheet_paper_size VALUE 41. "#EC NOTEXT
    CONSTANTS c_papersize_de_std_fanfold TYPE zexcel_sheet_paper_size VALUE 40. "#EC NOTEXT
    CONSTANTS c_papersize_dl_envelope TYPE zexcel_sheet_paper_size VALUE 27. "#EC NOTEXT
    CONSTANTS c_papersize_e TYPE zexcel_sheet_paper_size VALUE 26. "#EC NOTEXT
    CONSTANTS c_papersize_executive TYPE zexcel_sheet_paper_size VALUE 7. "#EC NOTEXT
    CONSTANTS c_papersize_folio TYPE zexcel_sheet_paper_size VALUE 14. "#EC NOTEXT
    CONSTANTS c_papersize_invite_envelope TYPE zexcel_sheet_paper_size VALUE 47. "#EC NOTEXT
    CONSTANTS c_papersize_iso_b4 TYPE zexcel_sheet_paper_size VALUE 42. "#EC NOTEXT
    CONSTANTS c_papersize_iso_b5_extra_paper TYPE zexcel_sheet_paper_size VALUE 63. "#EC NOTEXT
    CONSTANTS c_papersize_italy_envelope TYPE zexcel_sheet_paper_size VALUE 36. "#EC NOTEXT
    CONSTANTS c_papersize_jis_b5_tv_paper TYPE zexcel_sheet_paper_size VALUE 60. "#EC NOTEXT
    CONSTANTS c_papersize_jpn_dbl_postcard TYPE zexcel_sheet_paper_size VALUE 43. "#EC NOTEXT
    CONSTANTS c_papersize_ledger TYPE zexcel_sheet_paper_size VALUE 4. "#EC NOTEXT
    CONSTANTS c_papersize_legal TYPE zexcel_sheet_paper_size VALUE 5. "#EC NOTEXT
    CONSTANTS c_papersize_legal_extra_paper TYPE zexcel_sheet_paper_size VALUE 49. "#EC NOTEXT
    CONSTANTS c_papersize_letter TYPE zexcel_sheet_paper_size VALUE 1. "#EC NOTEXT
    CONSTANTS c_papersize_letter_extra_paper TYPE zexcel_sheet_paper_size VALUE 48. "#EC NOTEXT
    CONSTANTS c_papersize_letter_extv_paper TYPE zexcel_sheet_paper_size VALUE 54. "#EC NOTEXT
    CONSTANTS c_papersize_letter_plus_paper TYPE zexcel_sheet_paper_size VALUE 57. "#EC NOTEXT
    CONSTANTS c_papersize_letter_small TYPE zexcel_sheet_paper_size VALUE 2. "#EC NOTEXT
    CONSTANTS c_papersize_letter_tv_paper TYPE zexcel_sheet_paper_size VALUE 52. "#EC NOTEXT
    CONSTANTS c_papersize_monarch_envelope TYPE zexcel_sheet_paper_size VALUE 37. "#EC NOTEXT
    CONSTANTS c_papersize_no10_envelope TYPE zexcel_sheet_paper_size VALUE 20. "#EC NOTEXT
    CONSTANTS c_papersize_no11_envelope TYPE zexcel_sheet_paper_size VALUE 21. "#EC NOTEXT
    CONSTANTS c_papersize_no12_envelope TYPE zexcel_sheet_paper_size VALUE 22. "#EC NOTEXT
    CONSTANTS c_papersize_no14_envelope TYPE zexcel_sheet_paper_size VALUE 23. "#EC NOTEXT
    CONSTANTS c_papersize_no9_envelope TYPE zexcel_sheet_paper_size VALUE 19. "#EC NOTEXT
    CONSTANTS c_papersize_note TYPE zexcel_sheet_paper_size VALUE 18. "#EC NOTEXT
    CONSTANTS c_papersize_quarto TYPE zexcel_sheet_paper_size VALUE 15. "#EC NOTEXT
    CONSTANTS c_papersize_standard_1 TYPE zexcel_sheet_paper_size VALUE 16. "#EC NOTEXT
    CONSTANTS c_papersize_standard_2 TYPE zexcel_sheet_paper_size VALUE 17. "#EC NOTEXT
    CONSTANTS c_papersize_standard_paper_1 TYPE zexcel_sheet_paper_size VALUE 44. "#EC NOTEXT
    CONSTANTS c_papersize_standard_paper_2 TYPE zexcel_sheet_paper_size VALUE 45. "#EC NOTEXT
    CONSTANTS c_papersize_standard_paper_3 TYPE zexcel_sheet_paper_size VALUE 46. "#EC NOTEXT
    CONSTANTS c_papersize_statement TYPE zexcel_sheet_paper_size VALUE 6. "#EC NOTEXT
    CONSTANTS c_papersize_supera_a4_paper TYPE zexcel_sheet_paper_size VALUE 55. "#EC NOTEXT
    CONSTANTS c_papersize_superb_a3_paper TYPE zexcel_sheet_paper_size VALUE 56. "#EC NOTEXT
    CONSTANTS c_papersize_tabloid TYPE zexcel_sheet_paper_size VALUE 3. "#EC NOTEXT
    CONSTANTS c_papersize_tabl_extra_paper TYPE zexcel_sheet_paper_size VALUE 50. "#EC NOTEXT
    CONSTANTS c_papersize_us_std_fanfold TYPE zexcel_sheet_paper_size VALUE 39. "#EC NOTEXT
    CONSTANTS c_pe_blank TYPE string VALUE 'blank'.         "#EC NOTEXT
    CONSTANTS c_pe_dash TYPE string VALUE 'dash'.           "#EC NOTEXT
    CONSTANTS c_pe_displayed TYPE string VALUE 'displayed'. "#EC NOTEXT
    CONSTANTS c_pe_na TYPE string VALUE 'NA'.               "#EC NOTEXT
    DATA diff_oddeven_headerfooter TYPE flag .
    DATA draft TYPE flag .
    DATA errors TYPE string .
    DATA even_footer TYPE zexcel_s_worksheet_head_foot .
    DATA even_header TYPE zexcel_s_worksheet_head_foot .
    DATA first_page_number TYPE int2 .
    DATA fit_to_height TYPE int2 .
    DATA fit_to_page TYPE flag .
    DATA fit_to_width TYPE int2 .
    DATA horizontal_centered TYPE flag .
    DATA horizontal_dpi TYPE int2 .
    DATA margin_bottom TYPE zexcel_dec_8_2 .
    DATA margin_footer TYPE zexcel_dec_8_2 .
    DATA margin_header TYPE zexcel_dec_8_2 .
    DATA margin_left TYPE zexcel_dec_8_2 .
    DATA margin_right TYPE zexcel_dec_8_2 .
    DATA margin_top TYPE zexcel_dec_8_2 .
    DATA odd_footer TYPE zexcel_s_worksheet_head_foot .
    DATA odd_header TYPE zexcel_s_worksheet_head_foot .
    DATA orientation TYPE zexcel_sheet_orienatation .
    DATA page_order TYPE string .
    DATA paper_height TYPE string .
    DATA paper_size TYPE int2 .
    DATA paper_width TYPE string .
    DATA scale TYPE int2 .
    DATA use_first_page_num TYPE flag .
    DATA use_printer_defaults TYPE flag .
    DATA vertical_centered TYPE flag .
    DATA vertical_dpi TYPE int2 .

    METHODS constructor .
    METHODS set_page_margins
      IMPORTING
        !ip_bottom TYPE f OPTIONAL
        !ip_footer TYPE f OPTIONAL
        !ip_header TYPE f OPTIONAL
        !ip_left   TYPE f OPTIONAL
        !ip_right  TYPE f OPTIONAL
        !ip_top    TYPE f OPTIONAL
        !ip_unit   TYPE csequence DEFAULT 'in' .
    METHODS set_header_footer
      IMPORTING
        !ip_odd_header  TYPE zexcel_s_worksheet_head_foot OPTIONAL
        !ip_odd_footer  TYPE zexcel_s_worksheet_head_foot OPTIONAL
        !ip_even_header TYPE zexcel_s_worksheet_head_foot OPTIONAL
        !ip_even_footer TYPE zexcel_s_worksheet_head_foot OPTIONAL .
    METHODS get_header_footer_string
      EXPORTING
        !ep_odd_header  TYPE string
        !ep_odd_footer  TYPE string
        !ep_even_header TYPE string
        !ep_even_footer TYPE string .
    METHODS get_header_footer
      EXPORTING
        !ep_odd_header  TYPE zexcel_s_worksheet_head_foot
        !ep_odd_footer  TYPE zexcel_s_worksheet_head_foot
        !ep_even_header TYPE zexcel_s_worksheet_head_foot
        !ep_even_footer TYPE zexcel_s_worksheet_head_foot .