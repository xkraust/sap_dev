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
  PARAMETERS: p_file TYPE rlgrap-filename.             "Optional when p_bp is filled
  PARAMETERS: p_test AS CHECKBOX DEFAULT 'X'.          "Test run - no DB changes
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_bp  TYPE bu_partner.
  PARAMETERS: p_rfc TYPE rfcdest.
SELECTION-SCREEN END OF BLOCK b2.

*----------------------------------------------------------------------*
*  F4  H E L P   -   F I L E   P A T H
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.

  DATA(lo_view_f4) = NEW zmm_cl_bp_upd_view( ).
  p_file = lo_view_f4->get_file_path( ).

*----------------------------------------------------------------------*
*  S E L E C T I O N   S C R E E N   V A L I D A T I O N
*----------------------------------------------------------------------*
AT SELECTION-SCREEN.

  "-- Either a BP number or a file path must be provided -------------
  IF p_bp IS INITIAL AND p_file IS INITIAL.
    MESSAGE 'Please enter either a file path or a single BP number.'(e01) TYPE 'E'.
  ENDIF.


  IF p_bp IS NOT INITIAL AND p_rfc IS INITIAL.
    MESSAGE 'RFC destination is required when a BP number is entered.'(e02) TYPE 'E'.
  ENDIF.

*----------------------------------------------------------------------*
*  M A I N   P R O C E S S I N G
*----------------------------------------------------------------------*
START-OF-SELECTION.

  DATA(lo_model)     = NEW zmm_cl_bp_upd_model( ).
  DATA(lo_view)      = NEW zmm_cl_bp_upd_view( ).
  DATA(lo_presenter) = NEW zmm_cl_bp_upd_presenter( io_model = lo_model
                                                    io_view  = lo_view ).

  IF p_bp IS NOT INITIAL.
    lo_presenter->run_single_bp( iv_bp        = p_bp
                                 iv_rfc_dest  = p_rfc
                                 iv_test_mode = p_test ).
  ELSE.

    lo_presenter->run( iv_file      = p_file
                       iv_test_mode = p_test ).
  ENDIF.