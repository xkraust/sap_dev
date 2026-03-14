*&---------------------------------------------------------------------*
*& Report ZMM_UPDATE_EMAIL_ADDR
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zmm_update_email_addr.
*----------------------------------------------------------------------*
*  S E L E C T I O N   S C R E E N
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_file TYPE rlgrap-filename OBLIGATORY.
  PARAMETERS: p_test AS CHECKBOX DEFAULT 'X'.   "Test run - no DB changes
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
*  F4  H E L P   -   F I L E   P A T H
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.

  DATA(lo_view_f4) = NEW zmm_cl_bp_upd_view( ).
  p_file = lo_view_f4->get_file_path( ).

*----------------------------------------------------------------------*
*  M A I N   P R O C E S S I N G
*----------------------------------------------------------------------*
START-OF-SELECTION.

  DATA(lo_model)     = NEW zmm_cl_bp_upd_model( ).
  DATA(lo_view)      = NEW zmm_cl_bp_upd_view( ).
  DATA(lo_presenter) = NEW zmm_cl_bp_upd_presenter( io_model = lo_model
                                                    io_view  = lo_view ).

  lo_presenter->run( iv_file      = p_file
                     iv_test_mode = p_test ).