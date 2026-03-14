*&---------------------------------------------------------------------*
*& Global Class : ZCL_BP_UPD_PRESENTER
*& Description  : PRESENTER layer - orchestrates Model and View.
*&                Owns the run() method called from START-OF-SELECTION.
*&---------------------------------------------------------------------*
class ZMM_CL_BP_UPD_PRESENTER definition
  public
  final
  create public .

public section.

    "! Inject Model and View via constructor (dependency injection).
    "! @parameter io_model | ZCL_BP_UPD_MODEL instance
    "! @parameter io_view  | ZCL_BP_UPD_VIEW  instance
  methods CONSTRUCTOR
    importing
      !IO_MODEL type ref to ZMM_CL_BP_UPD_MODEL
      !IO_VIEW type ref to ZMM_CL_BP_UPD_VIEW .
    "! Main processing entry point.
    "! 1. Load Excel  2. Update BPs  3. Display ALV log
    "! @parameter iv_file      | Full local path to Excel file
    "! @parameter iv_test_mode | abap_true = simulate, no DB writes
  methods RUN
    importing
      !IV_FILE type LOCALFILE
      !IV_TEST_MODE type ABAP_BOOL default ABAP_TRUE .
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
  METHOD CONSTRUCTOR.

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
ENDCLASS.