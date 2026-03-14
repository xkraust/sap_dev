CLASS zmm_cl_bp_upd_view DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    METHODS get_file_path
      RETURNING
        VALUE(rv_path) TYPE rlgrap-filename .
    METHODS show_results
      IMPORTING
        !it_log       TYPE zmm_cl_bp_upd_model=>tt_errorlog
        !iv_test_mode TYPE abap_bool .
    METHODS show_message
      IMPORTING
        !iv_text TYPE string
        !iv_type TYPE symsgty DEFAULT 'I' .
  PRIVATE SECTION.

    "! Configure CL_SALV_TABLE column texts and widths
    METHODS configure_columns
      IMPORTING io_salv TYPE REF TO cl_salv_table.

    "! Set short / medium / long text and output length on a single column.
    "! Replaces the DEFINE macro previously used in configure_columns.
    "! @parameter io_cols      | Columns object from lo_salv->get_columns()
    "! @parameter iv_fieldname | Technical field name in ty_errorlog
    "! @parameter iv_short     | Short column header  (up to  8 chars)
    "! @parameter iv_medium    | Medium column header (up to 20 chars)
    "! @parameter iv_long      | Long column header   (up to 40 chars)
    "! @parameter iv_length    | Output length in characters
    METHODS set_column_texts
      IMPORTING io_cols      TYPE REF TO cl_salv_columns_table
                iv_fieldname TYPE lvc_fname
                iv_short     TYPE scrtext_s
                iv_medium    TYPE scrtext_m
                iv_long      TYPE scrtext_l
                iv_length    TYPE lvc_outlen.

ENDCLASS.



CLASS ZMM_CL_BP_UPD_VIEW IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_VIEW->CONFIGURE_COLUMNS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IO_SALV                        TYPE REF TO CL_SALV_TABLE
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD CONFIGURE_COLUMNS.

    DATA(lo_cols) = CAST cl_salv_columns_table( io_salv->get_columns( ) ).
    lo_cols->set_optimize( abap_true ).

    set_column_texts( io_cols = lo_cols  iv_fieldname = 'ROW'
      iv_short = 'Row'   iv_medium = 'Row No.'     iv_long = 'Row Number'        iv_length = 6  ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'TYPE'
      iv_short = 'Type'  iv_medium = 'Type'        iv_long = 'Msg Type'          iv_length = 6  ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'ACCID'
      iv_short = 'BP'    iv_medium = 'BP Number'   iv_long = 'Business Partner'  iv_length = 12 ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'ID'
      iv_short = 'Class' iv_medium = 'Msg Class'   iv_long = 'Message Class'     iv_length = 8  ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'NUMBER'
      iv_short = 'No.'   iv_medium = 'Msg No.'     iv_long = 'Message Number'    iv_length = 7  ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'MESSAGE'
      iv_short = 'BAPI'  iv_medium = 'BAPI Msg'    iv_long = 'BAPI Message'      iv_length = 60 ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'MSG'
      iv_short = 'Info'  iv_medium = 'Result'      iv_long = 'Result / Info Text' iv_length = 80 ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'LOG_MSG_NO'
      iv_short = 'LogNo' iv_medium = 'Log Msg No.' iv_long = 'Log Message Number' iv_length = 12 ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'MESSAGE_V1'
      iv_short = 'Var1'  iv_medium = 'Msg Var 1'   iv_long = 'Message Variable 1' iv_length = 40 ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'MESSAGE_V2'
      iv_short = 'Var2'  iv_medium = 'Msg Var 2'   iv_long = 'Message Variable 2' iv_length = 6  ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'MESSAGE_V3'
      iv_short = 'Var3'  iv_medium = 'Msg Var 3'   iv_long = 'Message Variable 3' iv_length = 25 ).
    set_column_texts( io_cols = lo_cols  iv_fieldname = 'MESSAGE_V4'
      iv_short = 'Var4'  iv_medium = 'Msg Var 4'   iv_long = 'Message Variable 4' iv_length = 10 ).
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_VIEW->GET_FILE_PATH
* +-------------------------------------------------------------------------------------------------+
* | [<-()] RV_PATH                        TYPE        RLGRAP-FILENAME
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD GET_FILE_PATH.

    DATA: lt_files  TYPE filetable,
          ls_file   TYPE file_table,
          lv_rc     TYPE i,
          lv_filter TYPE string.

    lv_filter = 'Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*'.

    CALL METHOD cl_gui_frontend_services=>file_open_dialog
      EXPORTING
        window_title            = 'Select Excel File for BP Update'
        default_extension       = 'xlsx'
        file_filter             = lv_filter
        multiselection          = abap_false
      CHANGING
        file_table              = lt_files
        rc                      = lv_rc
      EXCEPTIONS
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4
        OTHERS                  = 5.

    IF sy-subrc <> 0.
      show_message(
        iv_text = |File dialog error (rc={ sy-subrc })|
        iv_type = 'W' ).
      RETURN.
    ENDIF.

    IF lv_rc = 1.
      READ TABLE lt_files INDEX 1 INTO ls_file.
      rv_path = ls_file-filename.
    ENDIF.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_VIEW->SET_COLUMN_TEXTS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IO_COLS                        TYPE REF TO CL_SALV_COLUMNS_TABLE
* | [--->] IV_FIELDNAME                   TYPE        LVC_FNAME
* | [--->] IV_SHORT                       TYPE        SCRTEXT_S
* | [--->] IV_MEDIUM                      TYPE        SCRTEXT_M
* | [--->] IV_LONG                        TYPE        SCRTEXT_L
* | [--->] IV_LENGTH                      TYPE        LVC_OUTLEN
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD SET_COLUMN_TEXTS.

    TRY.
        DATA(lo_col) = CAST cl_salv_column_table(
                         io_cols->get_column( iv_fieldname ) ).
        lo_col->set_short_text(    iv_short  ).
        lo_col->set_medium_text(   iv_medium ).
        lo_col->set_long_text(     iv_long   ).
        lo_col->set_output_length( iv_length ).
      CATCH cx_salv_not_found.                            "#EC NO_HANDLER
    ENDTRY.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_VIEW->SHOW_MESSAGE
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_TEXT                        TYPE        STRING
* | [--->] IV_TYPE                        TYPE        SYMSGTY (default ='I')
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD SHOW_MESSAGE.

    MESSAGE iv_text TYPE iv_type.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_VIEW->SHOW_RESULTS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_LOG                         TYPE        ZMM_CL_BP_UPD_MODEL=>TT_ERRORLOG
* | [--->] IV_TEST_MODE                   TYPE        ABAP_BOOL
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD SHOW_RESULTS.

    DATA: lo_salv  TYPE REF TO cl_salv_table,
          lv_title TYPE lvc_title.

    "-- Work copy (factory needs CHANGING parameter) -----------------
    DATA(lt_log) = it_log.

    "-- Title --------------------------------------------------------
    lv_title = COND #( WHEN iv_test_mode = abap_true
                       THEN 'BP Update Result  -  TEST MODE (no changes written)'
                       ELSE 'BP Update Result  -  PRODUCTION run' ).

    "-- Create SALV instance -----------------------------------------
    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = lo_salv
          CHANGING  t_table      = lt_log ).
      CATCH cx_salv_msg INTO DATA(lx_msg).
        show_message( iv_text = lx_msg->get_text( ) iv_type = 'E' ).
        RETURN.
    ENDTRY.

    "-- Display settings (title + zebra) -----------------------------
    DATA(lo_disp) = lo_salv->get_display_settings( ).
    lo_disp->set_list_header( lv_title ).
    lo_disp->set_striped_pattern( cl_salv_display_settings=>true ).

    "-- All standard toolbar functions (export, sort, filter…) ------
    DATA(lo_funcs) = lo_salv->get_functions( ).
    lo_funcs->set_all( abap_true ).

    "-- Column configuration -----------------------------------------
    configure_columns( lo_salv ).

    "-- Sort: errors (E) first, then warnings (W), then success (S) --
    TRY.
        DATA(lo_sorts) = lo_salv->get_sorts( ).
        lo_sorts->add_sort(
          columnname = 'TYPE'
          sequence   = if_salv_c_sort=>sort_down ).
      CATCH cx_salv_not_found cx_salv_existing
            cx_salv_data_error.                         "#EC NO_HANDLER
    ENDTRY.

    "-- Selections ---------------------------------------------------
    DATA(lo_sel) = lo_salv->get_selections( ).
    lo_sel->set_selection_mode( if_salv_c_selection_mode=>row_column ).

    lo_salv->display( ).
  ENDMETHOD.
ENDCLASS.