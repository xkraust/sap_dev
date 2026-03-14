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
        consnumber      TYPE ad_consnum,
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
    METHODS fetch_bp_from_rfc
      IMPORTING
        !business_partner TYPE bu_partner
        !rfc_destination  TYPE rfcdest
      EXPORTING
        !address_data     TYPE tt_itab
        !error_message    TYPE string
      RETURNING
        VALUE(rv_ok)      TYPE sap_bool .
private section.

  constants LFA1_TABLE type TABNAME value 'LFA1' ##NO_TEXT.
  constants KNA1_TABLE type TABNAME value 'KNA1' ##NO_TEXT.
  constants ADR6_TABLE type TABNAME value 'ADR6' ##NO_TEXT.
  constants ADRT_TABLE type TABNAME value 'ADRT' ##NO_TEXT.

  methods COLLECT_BAPI_MSGS
    importing
      !IT_RETURN type BAPIRET2_T
    returning
      value(RV_TEXT) type BAPI_MSG .
  methods READ_REMOTE_TABLE
    importing
      !RFC_DESTINATION type RFCDEST
      !TABLE_NAME type TABNAME
      !WHERE_CLAUSE type RFC_DB_OPT
      !FIELD_NAMES type STRINGTAB optional
    exporting
      !RESULT_ROWS type ESH_T_CO_RFCRT_DATA
      !ERROR_MESSAGE type STRING
      !FIELDS_TAB type EHFNDT_DB_FIELDS
    returning
      value(RV_OK) type ABAP_BOOL .
  methods EXTRACT_FIELD_VALUE
    importing
      !DATA_LINE type ANY
      !FIELDS_TAB type EHFNDT_DB_FIELDS
      !FIELD_NAME type STRING
    returning
      value(RV_VALUE) type STRING .
ENDCLASS.



CLASS ZMM_CL_BP_UPD_MODEL IMPLEMENTATION.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->COLLECT_BAPI_MSGS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_RETURN                      TYPE        BAPIRET2_T
* | [<-()] RV_TEXT                        TYPE        BAPI_MSG
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD collect_bapi_msgs.

    LOOP AT it_return INTO DATA(ls_ret) WHERE type CA 'EAX'.
      rv_text = COND #( WHEN rv_text IS INITIAL
                        THEN ls_ret-message
                        ELSE |{ rv_text } { ls_ret-message }| ).
    ENDLOOP.
  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->EXTRACT_FIELD_VALUE
* +-------------------------------------------------------------------------------------------------+
* | [--->] DATA_LINE                      TYPE        ANY
* | [--->] FIELDS_TAB                     TYPE        EHFNDT_DB_FIELDS
* | [--->] FIELD_NAME                     TYPE        STRING
* | [<-()] RV_VALUE                       TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD extract_field_value.

    DATA(lv_line) = CONV string( data_line ).
    DATA(lv_position) = 0.
    DATA(lt_parts) = VALUE stringtab( ).

    SPLIT lv_line AT '|' INTO TABLE lt_parts.

    LOOP AT fields_tab ASSIGNING FIELD-SYMBOL(<field>).
      lv_position += 1.
      DATA(current_name) = CONV string( <field>-fieldname ).
      IF current_name = field_name.
        TRY.
            rv_value = condense( lt_parts[ lv_position ] ).
          CATCH cx_sy_itab_line_not_found.
            rv_value = ''.
        ENDTRY.
        RETURN.
      ENDIF.
    ENDLOOP.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_MODEL->FETCH_BP_FROM_RFC
* +-------------------------------------------------------------------------------------------------+
* | [--->] BUSINESS_PARTNER               TYPE        BU_PARTNER
* | [--->] RFC_DESTINATION                TYPE        RFCDEST
* | [<---] ADDRESS_DATA                   TYPE        TT_ITAB
* | [<---] ERROR_MESSAGE                  TYPE        STRING
* | [<-()] RV_OK                          TYPE        SAP_BOOL
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD fetch_bp_from_rfc.
    DATA lfa1_rows   TYPE esh_t_co_rfcrt_data.
    DATA lfa1_fields TYPE ehfndt_db_fields.
    DATA kna1_rows   TYPE esh_t_co_rfcrt_data.
    DATA kna1_fields TYPE ehfndt_db_fields.
    DATA adr6_rows   TYPE esh_t_co_rfcrt_data.
    DATA adr6_fields TYPE ehfndt_db_fields.
    DATA adrt_rows   TYPE esh_t_co_rfcrt_data.
    DATA adrt_fields TYPE ehfndt_db_fields.
    DATA address_tmp TYPE tt_itab.

    CLEAR address_data.
    rv_ok = abap_false.

    "-- Pad BP number with leading zeros (10 chars) for WHERE clause --------
    DATA(bp_padded) = CONV bu_partner( |{ business_partner ALPHA = IN }| ).

    "-- 1. Look up address number: check LFA1 (vendor) first ----------------
    DATA(where_lfa1)    = CONV rfc_db_opt( |LIFNR EQ '{ bp_padded }'| ).
    DATA(fields_lfa1)   = VALUE stringtab( ( |LIFNR| ) ( |ADRNR| ) ).
    DATA(found_in_lfa1) = read_remote_table( EXPORTING rfc_destination = rfc_destination
                                                       table_name      = lfa1_table
                                                       where_clause    = where_lfa1
                                                       field_names     = fields_lfa1
                                             IMPORTING result_rows     = lfa1_rows
                                                       fields_tab      = lfa1_fields
                                                       error_message   = error_message ).

    DATA address_number TYPE adrnr.

    IF found_in_lfa1 = abap_true AND lines( lfa1_rows ) > 0.
      address_number = extract_field_value( data_line  = lfa1_rows[ 1 ]
                                            fields_tab = lfa1_fields
                                            field_name = 'ADRNR' ).
    ENDIF.

    "-- 2. Fall back to KNA1 (customer) if not found in LFA1 ----------------
    IF address_number IS INITIAL.
      DATA(where_kna1)  = CONV rfc_db_opt( |KUNNR EQ '{ bp_padded }'| ).
      DATA(fields_kna1) = VALUE stringtab( ( |KUNNR| ) ( |ADRNR| ) ).


      DATA(found_in_kna1) = read_remote_table( EXPORTING rfc_destination = rfc_destination
                                                         table_name      = kna1_table
                                                         where_clause    = where_kna1
                                                         field_names     = fields_kna1
                                               IMPORTING result_rows     = kna1_rows
                                                         fields_tab      = kna1_fields
                                                         error_message   = error_message ).

      IF found_in_kna1 = abap_true AND lines( kna1_rows ) > 0.
        address_number = extract_field_value( data_line  = kna1_rows[ 1 ]
                                              fields_tab = kna1_fields
                                              field_name = 'ADRNR' ).
      ENDIF.
    ENDIF.

    IF address_number IS INITIAL.
      error_message = |BP { business_partner }: not found in LFA1 or KNA1 on remote system|.
      RETURN.
    ENDIF.

    "-- 3. Read e-mail from ADR6 ---------------------------------------------
    DATA(where_adr6)  = CONV rfc_db_opt( |ADDRNUMBER EQ '{ address_number }'| ).
    DATA(fields_adr6) = VALUE stringtab( ( |ADDRNUMBER| ) ( |SMTP_ADDR| ) ( |FLGDEFAULT| ) ( |CONSNUMBER| ) ).


    DATA(adr6_ok) = read_remote_table( EXPORTING rfc_destination = rfc_destination
                                                 table_name      = adr6_table
                                                 where_clause    = where_adr6
                                                 field_names     = fields_adr6
                                       IMPORTING result_rows     = adr6_rows
                                                 fields_tab      = adr6_fields
                                                 error_message   = error_message ).

    IF adr6_ok = abap_true AND lines( adr6_rows ) > 0.
      LOOP AT adr6_rows ASSIGNING FIELD-SYMBOL(<fs_addr6_row>).
        APPEND INITIAL LINE TO address_data ASSIGNING FIELD-SYMBOL(<fs_address_data>).
        <fs_address_data>-email = extract_field_value( data_line  = <fs_addr6_row>
                                                       fields_tab = adr6_fields
                                                       field_name = 'SMTP_ADDR' ).
        <fs_address_data>-std_no = extract_field_value( data_line  = <fs_addr6_row>
                                                       fields_tab = adr6_fields
                                                       field_name = 'FLGDEFAULT' ).
        <fs_address_data>-consnumber = extract_field_value( data_line  = <fs_addr6_row>
                                                            fields_tab = adr6_fields
                                                            field_name = 'CONSNUMBER' ).

        <fs_address_data>-email_x = abap_true.
        <fs_address_data>-businesspartner = bp_padded.
      ENDLOOP.

    ENDIF.

    "-- 4. Read address note from ADRT ---------------------------------------
    DATA(where_adrt)  = CONV rfc_db_opt( |ADDRNUMBER EQ '{ address_number }'| ).
    DATA(fields_adrt) = VALUE stringtab( ( |ADDRNUMBER| ) ( |REMARK| ) ( |LANGU| ) ( |CONSNUMBER| ) ).


    DATA(adrt_ok) = read_remote_table( EXPORTING rfc_destination = rfc_destination
                                                 table_name      = adrt_table
                                                 where_clause    = where_adrt
                                                 field_names     = fields_adrt
                                       IMPORTING result_rows     = adrt_rows
                                                 fields_tab      = adrt_fields
                                                 error_message   = error_message ).

    IF adrt_ok = abap_true AND lines( adrt_rows ) > 0.
      LOOP AT adrt_rows ASSIGNING FIELD-SYMBOL(<fs_addrt_row>).
        APPEND INITIAL LINE TO address_tmp ASSIGNING <fs_address_data>.
        <fs_address_data>-notes = extract_field_value( data_line  = <fs_addrt_row>
                                                       fields_tab = adrt_fields
                                                       field_name = 'REMARK' ).
        <fs_address_data>-language = extract_field_value( data_line  = <fs_addrt_row>
                                                          fields_tab = adrt_fields
                                                          field_name = 'LANGU' ).
        <fs_address_data>-consnumber = extract_field_value( data_line  = <fs_addrt_row>
                                                            fields_tab = adrt_fields
                                                            field_name = 'CONSNUMBER' ).
        <fs_address_data>-notes_x  = abap_true.
      ENDLOOP.
    ENDIF.

    "-- 5. Merge emails and remarks -----------------------------------------
    LOOP AT address_data ASSIGNING <fs_address_data>.
      TRY .
          <fs_address_data>-notes    = address_tmp[ consnumber = <fs_address_data>-consnumber  ]-notes.
          <fs_address_data>-language = address_tmp[ consnumber = <fs_address_data>-consnumber  ]-language.
          <fs_address_data>-notes_x  = COND #( WHEN <fs_address_data>-notes IS NOT INITIAL THEN abap_true ELSE abap_false ).
        CATCH cx_sy_itab_line_not_found.
          CONTINUE.
      ENDTRY.
    ENDLOOP.

    "-- 6. Guard: at least email or note must be present --------------------
    IF address_data IS INITIAL.
      error_message = |BP { business_partner }: no e-mail or address note found in remote ADR6/ADRT|.
      RETURN.
    ENDIF.

    rv_ok = abap_true.

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
      ev_error_msg = |Excel upload failed (ALSM_EXCEL_TO_INTERNAL_TABLE rc={ sy-subrc })|.
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
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->READ_REMOTE_TABLE
* +-------------------------------------------------------------------------------------------------+
* | [--->] RFC_DESTINATION                TYPE        RFCDEST
* | [--->] TABLE_NAME                     TYPE        TABNAME
* | [--->] WHERE_CLAUSE                   TYPE        RFC_DB_OPT
* | [--->] FIELD_NAMES                    TYPE        STRINGTAB(optional)
* | [<---] RESULT_ROWS                    TYPE        ESH_T_CO_RFCRT_DATA
* | [<---] ERROR_MESSAGE                  TYPE        STRING
* | [<---] FIELDS_TAB                     TYPE        EHFNDT_DB_FIELDS
* | [<-()] RV_OK                          TYPE        ABAP_BOOL
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD read_remote_table.

    DATA where_tab TYPE STANDARD TABLE OF rfc_db_opt.
    where_tab = VALUE #( ( where_clause ) ).


    DATA fields TYPE STANDARD TABLE OF rfc_db_fld.
    LOOP AT field_names INTO DATA(field_name).
      APPEND VALUE #( fieldname = field_name ) TO fields.
    ENDLOOP.

    DATA data_tab  TYPE STANDARD TABLE OF tab512 WITH EMPTY KEY.
    DATA field_tab TYPE STANDARD TABLE OF tab512 WITH EMPTY KEY.

    CALL FUNCTION 'RFC_READ_TABLE'
      DESTINATION rfc_destination
      EXPORTING
        query_table          = table_name
        delimiter            = '|'
      TABLES
        options              = where_tab
        fields               = fields
        data                 = data_tab
      EXCEPTIONS
        table_not_available  = 1
        table_without_data   = 2
        option_not_valid     = 3
        field_not_valid      = 4
        not_authorized       = 5
        data_buffer_exceeded = 6
        OTHERS               = 7.

    IF sy-subrc <> 0.
      error_message = |RFC_READ_TABLE on { table_name } failed (rc={ sy-subrc })|.
      rv_ok = abap_false.
      RETURN.
    ENDIF.

    result_rows = data_tab.
    rv_ok       = abap_true.
    fields_tab  = fields.

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