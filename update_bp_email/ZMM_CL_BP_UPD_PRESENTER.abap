*&---------------------------------------------------------------------*
*& Global Class : ZCL_BP_UPD_PRESENTER
*&---------------------------------------------------------------------*
CLASS zmm_cl_bp_upd_presenter DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.


    METHODS constructor
      IMPORTING
        !io_model TYPE REF TO zmm_cl_bp_upd_model
        !io_view  TYPE REF TO zmm_cl_bp_upd_view .
    METHODS run
      IMPORTING
        !iv_file      TYPE localfile
        !iv_test_mode TYPE abap_bool DEFAULT abap_true .

    METHODS run_single_bp
      IMPORTING
        !iv_bp        TYPE bu_partner
        !iv_rfc_dest  TYPE rfcdest
        !iv_test_mode TYPE abap_bool DEFAULT abap_true.

  PRIVATE SECTION.

    DATA: mo_model TYPE REF TO zmm_cl_bp_upd_model,
          mo_view  TYPE REF TO zmm_cl_bp_upd_view.
ENDCLASS.



CLASS ZMM_CL_BP_UPD_PRESENTER IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_PRESENTER->CONSTRUCTOR
* +-------------------------------------------------------------------------------------------------+
* | [--->] IO_MODEL                       TYPE REF TO ZMM_CL_BP_UPD_MODEL
* | [--->] IO_VIEW                        TYPE REF TO ZMM_CL_BP_UPD_VIEW
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD constructor.

    mo_model = io_model.
    mo_view  = io_view.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_PRESENTER->RUN
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_FILE                        TYPE        LOCALFILE
* | [--->] IV_TEST_MODE                   TYPE        ABAP_BOOL (default =ABAP_TRUE)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD run.

    DATA: lt_data   TYPE zmm_cl_bp_upd_model=>tt_itab,
          lt_log    TYPE zmm_cl_bp_upd_model=>tt_errorlog,
          lv_errmsg TYPE string,
          lv_ok     TYPE abap_bool.

    "-- 1. Load and parse Excel file ---------------------------------
    lv_ok = mo_model->load_excel( EXPORTING iv_file      = iv_file
                                  IMPORTING et_data      = lt_data
                                            ev_error_msg = lv_errmsg ).

    IF lv_ok = abap_false.
      mo_view->show_message( iv_text = lv_errmsg  iv_type = 'E' ).
      RETURN.
    ENDIF.

    IF lt_data IS INITIAL.
      mo_view->show_message( iv_text = 'No rows loaded from Excel file - processing stopped.'
                             iv_type = 'W' ).
      RETURN.
    ENDIF.

    "-- 2. Update Business Partners ----------------------------------
    lt_log = mo_model->update_partners( it_data      = lt_data
                                        iv_test_mode = iv_test_mode ).

    "-- 3. Display result in ALV -------------------------------------
    mo_view->show_results( it_log       = lt_log
                           iv_test_mode = iv_test_mode ).

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_PRESENTER->RUN_SINGLE_BP
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_BP                          TYPE        BU_PARTNER
* | [--->] IV_RFC_DEST                    TYPE        RFCDEST
* | [--->] IV_TEST_MODE                   TYPE        ABAP_BOOL (default =ABAP_TRUE)
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD run_single_bp.

    "-- 1. Read address data from the remote system via RFC_READ_TABLE -------
    DATA lt_address_data  TYPE zmm_cl_bp_upd_model=>tt_itab.
    DATA lv_fetch_error   TYPE string.

    DATA(fetch_ok) = mo_model->fetch_bp_from_rfc( EXPORTING business_partner = iv_bp
                                                            rfc_destination  = iv_rfc_dest
                                                  IMPORTING address_data     = lt_address_data
                                                            error_message    = lv_fetch_error ).

    IF fetch_ok = abap_false.
      mo_view->show_message( iv_text = lv_fetch_error  iv_type = 'E' ).
      RETURN.
    ENDIF.

    DATA(update_log) = mo_model->update_partners( it_data      = lt_address_data
                                                  iv_test_mode = iv_test_mode ).

    "-- 2. Display result in ALV --------------------------------------------
    mo_view->show_results( it_log       = update_log
                           iv_test_mode = iv_test_mode ).

  ENDMETHOD.
ENDCLASS.