CLASS zmm_cl_bp_upd_model DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ty_itab,
        businesspartner TYPE bu_partner,
        email_x         TYPE bapiadsmtx-e_mail,       "update flag email
        notes_x         TYPE bapicomrex-comm_notes,   "update flag notes
        email           TYPE ad_smtpadr,
        notes           TYPE bapicomrem-comm_notes,
        std_no          TYPE ad_flgdfnr,
        language        TYPE syst_langu,
      END OF ty_itab .
    TYPES:
      tt_itab TYPE STANDARD TABLE OF ty_itab WITH EMPTY KEY .
    TYPES:
      "Result / log row shown in ALV
      BEGIN OF ty_errorlog,
        row        TYPE i,
        type       TYPE bapiret2-type,
        accid      TYPE c LENGTH 10,              "BP number display
        id         TYPE bapiret2-id,
        number     TYPE bapiret2-number,
        message    TYPE bapiret2-message,
        msg        TYPE bapi_msg,                 "free-text result
        log_msg_no TYPE bapiret2-log_msg_no,
        message_v1 TYPE bapiret2-message_v1,
        message_v2 TYPE bapiret2-message_v2,
        message_v3 TYPE bapiret2-message_v3,
        message_v4 TYPE bapiret2-message_v4,
      END OF ty_errorlog .
    TYPES:
      tt_errorlog TYPE STANDARD TABLE OF ty_errorlog WITH EMPTY KEY .

    METHODS load_excel
      IMPORTING
        !iv_file      TYPE localfile
      EXPORTING
        !et_data      TYPE tt_itab
        !ev_error_msg TYPE string
      RETURNING
        VALUE(rv_ok)  TYPE abap_bool .
    METHODS update_partners
      IMPORTING
        !it_data      TYPE tt_itab
        !iv_test_mode TYPE abap_bool DEFAULT abap_true
      RETURNING
        VALUE(rt_log) TYPE tt_errorlog .
  PRIVATE SECTION.

    METHODS:

      collect_bapi_msgs
        IMPORTING it_return      TYPE bapiret2_t
        RETURNING VALUE(rv_text) TYPE bapi_msg.

ENDCLASS.



CLASS ZMM_CL_BP_UPD_MODEL IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->COLLECT_BAPI_MSGS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_RETURN                      TYPE        BAPIRET2_T
* | [<-()] RV_TEXT                        TYPE        BAPI_MSG
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD COLLECT_BAPI_MSGS.

    LOOP AT it_return INTO DATA(ls_ret) WHERE type CA 'EAX'.
      rv_text = COND #( WHEN rv_text IS INITIAL
                        THEN ls_ret-message
                        ELSE |{ rv_text } { ls_ret-message }| ).
    ENDLOOP.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_MODEL->LOAD_EXCEL
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_FILE                        TYPE        LOCALFILE
* | [<---] ET_DATA                        TYPE        TT_ITAB
* | [<---] EV_ERROR_MSG                   TYPE        STRING
* | [<-()] RV_OK                          TYPE        ABAP_BOOL
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD load_excel.

    DATA lt_excel TYPE STANDARD TABLE OF alsmex_tabline.
    FIELD-SYMBOLS <fs_data> TYPE ty_itab.

    rv_ok = abap_false.
    CLEAR et_data.

    CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
      EXPORTING
        filename                = iv_file
        i_begin_col             = 1
        i_begin_row             = 1
        i_end_col               = 50
        i_end_row               = 9999
      TABLES
        intern                  = lt_excel
      EXCEPTIONS
        inconsistent_parameters = 1
        upload_ole              = 2
        OTHERS                  = 3.

    IF sy-subrc <> 0.
      ev_error_msg = |Excel upload failed (TEXT_CONVERT_XLS_TO_SAP rc={ sy-subrc })|.
      RETURN.
    ENDIF.

    LOOP AT lt_excel ASSIGNING FIELD-SYMBOL(<fs_excel>) WHERE row <> '0001'.
      AT NEW row.
        APPEND INITIAL LINE TO et_data ASSIGNING <fs_data>.
        <fs_data>-businesspartner = <fs_excel>-value.
      ENDAT.
      CASE <fs_excel>-col.
        WHEN '0002'.
          <fs_data>-email_x = <fs_excel>-value.
        WHEN '0003'.
          <fs_data>-notes_x = <fs_excel>-value.
        WHEN '0004'.
          <fs_data>-email = <fs_excel>-value.
        WHEN '0005'.
          <fs_data>-notes = <fs_excel>-value.
        WHEN '0005'.
          <fs_data>-std_no = <fs_excel>-value.
      ENDCASE.
    ENDLOOP.

    IF et_data IS INITIAL.
      ev_error_msg = 'No data rows found after removing header row'.
      RETURN.
    ENDIF.

    rv_ok = abap_true.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_MODEL->UPDATE_PARTNERS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_DATA                        TYPE        TT_ITAB
* | [--->] IV_TEST_MODE                   TYPE        ABAP_BOOL (default =ABAP_TRUE)
* | [<-()] RT_LOG                         TYPE        TT_ERRORLOG
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD update_partners.

    DATA: lv_addrnumber TYPE but020-addrnumber,
          lt_adsmtp     TYPE STANDARD TABLE OF bapiadsmtp,
          ls_adsmtp     TYPE bapiadsmtp,
          lt_adsmt_x    TYPE STANDARD TABLE OF bapiadsmtx,
          ls_adsmt_x    TYPE bapiadsmtx,
          lt_comrem     TYPE STANDARD TABLE OF bapicomrem,
          ls_comrem     TYPE bapicomrem,
          lt_comre_x    TYPE STANDARD TABLE OF bapicomrex,
          ls_comre_x    TYPE bapicomrex,
          lt_return     TYPE bapiret2_t,
          ls_log        TYPE ty_errorlog,
          lv_row        TYPE i.

    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<fs_data>).
      lv_row += 1.
      CLEAR: ls_log,
             lt_adsmtp, ls_adsmtp, lt_adsmt_x, ls_adsmt_x,
             lt_comrem, ls_comrem, lt_comre_x, ls_comre_x,
             lt_return, lv_addrnumber.

      ls_log = VALUE #( row   = lv_row
                        accid = <fs_data>-businesspartner ).

      "-- Validate: BP number must be filled -------------------------
      IF <fs_data>-businesspartner IS INITIAL.
        rt_log = VALUE #( BASE rt_log ( row = ls_log-row accid = ls_log-accid type = 'E' msg  = |BP number is empty - row skipped| ) ).
        CONTINUE.
      ENDIF.

      "-- Validate: at least one payload field -----------------------
      IF <fs_data>-email IS INITIAL AND <fs_data>-notes IS INITIAL.
        rt_log = VALUE #( BASE rt_log ( row = ls_log-row accid = ls_log-accid type = 'W' msg  = |BP { <fs_data>-businesspartner }: email and note both empty - skipped| ) ).
        CONTINUE.
      ENDIF.

      DATA(lv_bp) = CONV bu_partner( |{ <fs_data>-businesspartner ALPHA = IN }| ).

      "-- Build SMTP (email) structures ------------------------------
      IF <fs_data>-email IS NOT INITIAL.
        lt_adsmtp  = VALUE #( ( e_mail = <fs_data>-email consnumber = '001' std_no = <fs_data>-std_no ) ).
        lt_adsmt_x = VALUE #( ( e_mail = COND #( WHEN <fs_data>-email_x IS INITIAL THEN 'X' ELSE <fs_data>-email_x ) consnumber = 'X' updateflag = 'I' ) ).
      ENDIF.

      "-- Build COMM_NOTES (address note) structures -----------------
      IF <fs_data>-notes IS NOT INITIAL.
        lt_comrem  = VALUE #( ( comm_type = 'INT' langu = <fs_data>-language comm_notes  = <fs_data>-notes ) ).
        lt_comre_x = VALUE #( ( comm_notes = COND #( WHEN <fs_data>-notes_x IS INITIAL THEN 'X' ELSE <fs_data>-notes_x ) updateflag = 'I' ) ).
      ENDIF.

      "-- Call BAPI_BUPA_ADDRESS_CHANGE (skip DB write in test mode) -
      CALL FUNCTION 'BAPI_BUPA_ADDRESS_CHANGE'
        EXPORTING
          businesspartner = lv_bp
        TABLES
          bapiadsmtp      = lt_adsmtp
          bapicomrem      = lt_comrem
          bapiadsmt_x     = lt_adsmt_x
          bapicomre_x     = lt_comre_x
          return          = lt_return.

      "-- Evaluate result --------------------------------------------
      IF NOT line_exists( lt_return[ type = 'E' ] ) AND NOT line_exists( lt_return[ type = 'A' ] ).
        IF iv_test_mode = abap_false.
          "-- Success ---------------------------------------------------
          CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
            EXPORTING
              wait = abap_true.
          rt_log = VALUE #( BASE rt_log ( row = ls_log-row accid = ls_log-accid type = 'S'
                                  msg = |OK - Email: { <fs_data>-email } / Note: { <fs_data>-notes } / BP: { <fs_data>-businesspartner }| ) ).
        ELSE.
          CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
          rt_log = VALUE #( BASE rt_log ( row = ls_log-row accid = ls_log-accid type = 'S'
                                        msg = |TEST - Email: { <fs_data>-email } / Note: { <fs_data>-notes } / BP: { <fs_data>-businesspartner } - no changes written| ) ).
        ENDIF.
      ELSE.
        "-- Error -----------------------------------------------------
        CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

        ls_log = VALUE #( type = 'E' msg = collect_bapi_msgs( lt_return ) ).

        "Copy first error detail into log fields for ALV columns
        TRY .
            DATA(ls_ret) = lt_return[ type = 'E' ].

            rt_log = VALUE #( BASE rt_log ( row        = lv_row
                                            accid      = lv_bp
                                            id         = ls_ret-id
                                            type       = 'E'
                                            number     = ls_ret-number
                                            message    = ls_ret-message
                                            log_msg_no = ls_ret-log_msg_no
                                            message_v1 = ls_ret-message_v1
                                            message_v2 = ls_ret-message_v2
                                            message_v3 = ls_ret-message_v3
                                            message_v4 = ls_ret-message_v4 ) ).
          CATCH cx_sy_itab_line_not_found.
        ENDTRY.

      ENDIF.
    ENDLOOP.

  ENDMETHOD.
ENDCLASS.