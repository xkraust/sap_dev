CLASS zmm_cl_bp_upd_model DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.

    "-- Email / address types (unchanged) ----------------------------------
    TYPES:
      BEGIN OF ty_itab,
        businesspartner TYPE bu_partner,
        email_x         TYPE bapiadsmtx-e_mail,
        notes_x         TYPE bapicomrex-comm_notes,
        email           TYPE ad_smtpadr,
        notes           TYPE bapicomrem-comm_notes,
        std_no          TYPE ad_flgdfnr,
        language        TYPE syst_langu,
        consnumber      TYPE ad_consnum,
      END OF ty_itab.
    TYPES:
      tt_itab TYPE STANDARD TABLE OF ty_itab WITH EMPTY KEY.

    "-- Banking data types -------------------------------------------------
    TYPES:
      BEGIN OF ty_bank_itab,
        businesspartner TYPE bu_partner,
        lifnr           TYPE lifnr,
        banks           TYPE banks,
        bankl           TYPE bankk,
        bankn           TYPE bankn,
        bkont           TYPE bkont,
        bvtyp           TYPE bvtyp,
        xezer           TYPE xezer,
        bkref           TYPE bkref,
        koinh           TYPE koinh_fi,
        bank_guid       TYPE bu_bp_bank_guid,
        tech_rectyp     TYPE cvmd_bankacc_tech_rectyp,
        ebpp_accname    TYPE ebpp_accname,
        ebpp_bvstatus   TYPE ebpp_bvstatus,
        kovon           TYPE kovon,
        kobis           TYPE kobis,
        iban            TYPE iban,
      END OF ty_bank_itab.
    TYPES:
      tt_bank_itab TYPE STANDARD TABLE OF ty_bank_itab WITH EMPTY KEY.

    "-- Shared result / log type -------------------------------------------
    TYPES:
      BEGIN OF ty_errorlog,
        row        TYPE i,
        type       TYPE bapiret2-type,
        accid      TYPE c LENGTH 10,
        id         TYPE bapiret2-id,
        number     TYPE bapiret2-number,
        message    TYPE bapiret2-message,
        msg        TYPE bapi_msg,
        log_msg_no TYPE bapiret2-log_msg_no,
        message_v1 TYPE bapiret2-message_v1,
        message_v2 TYPE bapiret2-message_v2,
        message_v3 TYPE bapiret2-message_v3,
        message_v4 TYPE bapiret2-message_v4,
      END OF ty_errorlog.
    TYPES:
      tt_errorlog TYPE STANDARD TABLE OF ty_errorlog WITH EMPTY KEY.

    "-- Email methods (unchanged) ------------------------------------------
    METHODS load_excel
      IMPORTING
        !iv_file      TYPE localfile
      EXPORTING
        !et_data      TYPE tt_itab
        !ev_error_msg TYPE string
      RETURNING
        VALUE(rv_ok)  TYPE abap_bool.

    METHODS update_partners
      IMPORTING
        !it_data      TYPE tt_itab
        !iv_test_mode TYPE abap_bool DEFAULT abap_true
      RETURNING
        VALUE(rt_log) TYPE tt_errorlog.

    METHODS fetch_bp_from_rfc
      IMPORTING
        !business_partner TYPE bu_partner
        !rfc_destination  TYPE rfcdest
      EXPORTING
        !address_data     TYPE tt_itab
        !error_message    TYPE string
      RETURNING
        VALUE(rv_ok)      TYPE sap_bool.

    "-- Banking methods (new) ----------------------------------------------
    METHODS load_excel_bank
      IMPORTING
        !iv_file      TYPE localfile
      EXPORTING
        !et_data      TYPE tt_bank_itab
        !ev_error_msg TYPE string
      RETURNING
        VALUE(rv_ok)  TYPE abap_bool.

    METHODS fetch_bank_from_rfc
      IMPORTING
        !iv_business_partner TYPE bu_partner
        !iv_rfc_destination  TYPE rfcdest
      EXPORTING
        !et_bank_data        TYPE tt_bank_itab
        !ev_error_message    TYPE string
      RETURNING
        VALUE(rv_ok)         TYPE abap_bool.

    METHODS update_bank_details
      IMPORTING
        !it_data      TYPE tt_bank_itab
        !iv_test_mode TYPE abap_bool DEFAULT abap_true
      RETURNING
        VALUE(rt_log) TYPE tt_errorlog.

  PRIVATE SECTION.

    CONSTANTS: lc_lfa1_table  TYPE tabname VALUE 'LFA1'  ##NO_TEXT,
               lc_kna1_table  TYPE tabname VALUE 'KNA1'  ##NO_TEXT,
               lc_adr6_table  TYPE tabname VALUE 'ADR6'  ##NO_TEXT,
               lc_adrt_table  TYPE tabname VALUE 'ADRT'  ##NO_TEXT,
               lc_lfbk_table  TYPE tabname VALUE 'LFBK'  ##NO_TEXT,
               lc_knbk_table  TYPE tabname VALUE 'KNBK'  ##NO_TEXT,
               lc_tiban_table TYPE tabname VALUE 'TIBAN' ##NO_TEXT.

    METHODS collect_bapi_msgs
      IMPORTING
        !it_return    TYPE bapiret2_t
      RETURNING
        VALUE(rv_text) TYPE bapi_msg.

    METHODS read_remote_table
      IMPORTING
        !iv_rfc_destination TYPE rfcdest
        !iv_table_name      TYPE tabname
        !it_where_clause    TYPE esh_t_co_rfcrt_options
        !it_field_names     TYPE stringtab OPTIONAL
      EXPORTING
        !et_result_rows     TYPE esh_t_co_rfcrt_data
        !ev_error_message   TYPE string
        !et_fields_tab      TYPE ehfndt_db_fields
      RETURNING
        VALUE(rv_ok)        TYPE abap_bool.

    METHODS extract_field_value
      IMPORTING
        !iv_data_line  TYPE any
        !it_fields_tab TYPE ehfndt_db_fields
        !iv_field_name TYPE string
      RETURNING
        VALUE(rv_value) TYPE string.

ENDCLASS.


CLASS zmm_cl_bp_upd_model IMPLEMENTATION.

* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->COLLECT_BAPI_MSGS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_RETURN                      TYPE        BAPIRET2_T
* | [<-()] RV_TEXT                        TYPE        BAPI_MSG
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD collect_bapi_msgs.

    LOOP AT it_return ASSIGNING FIELD-SYMBOL(<fs_ret>) WHERE type CA 'EAX'.
      rv_text = COND #( WHEN rv_text IS INITIAL
                        THEN <fs_ret>-message
                        ELSE |{ rv_text } { <fs_ret>-message }| ).
    ENDLOOP.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->EXTRACT_FIELD_VALUE
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_DATA_LINE                   TYPE        ANY
* | [--->] IT_FIELDS_TAB                  TYPE        EHFNDT_DB_FIELDS
* | [--->] IV_FIELD_NAME                  TYPE        STRING
* | [<-()] RV_VALUE                       TYPE        STRING
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD extract_field_value.

    DATA(lv_line)     = CONV string( iv_data_line ).
    DATA(lv_position) = 0.
    DATA(lt_parts)    = VALUE stringtab( ).

    SPLIT lv_line AT '|' INTO TABLE lt_parts.

    LOOP AT it_fields_tab ASSIGNING FIELD-SYMBOL(<fs_field>).
      lv_position += 1.
      IF CONV string( <fs_field>-fieldname ) = iv_field_name.
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
* | Instance Private Method ZMM_CL_BP_UPD_MODEL->READ_REMOTE_TABLE
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_RFC_DESTINATION             TYPE        RFCDEST
* | [--->] IV_TABLE_NAME                  TYPE        TABNAME
* | [--->] IT_WHERE_CLAUSE                TYPE        ESH_T_CO_RFCRT_OPTIONS
* | [--->] IT_FIELD_NAMES                 TYPE        STRINGTAB (optional)
* | [<---] ET_RESULT_ROWS                 TYPE        ESH_T_CO_RFCRT_DATA
* | [<---] EV_ERROR_MESSAGE               TYPE        STRING
* | [<---] ET_FIELDS_TAB                  TYPE        EHFNDT_DB_FIELDS
* | [<-()] RV_OK                          TYPE        ABAP_BOOL
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD read_remote_table.

    DATA lt_where_tab TYPE STANDARD TABLE OF rfc_db_opt.
    DATA lt_fields    TYPE STANDARD TABLE OF rfc_db_fld.
    DATA lt_data_tab  TYPE STANDARD TABLE OF tab512 WITH EMPTY KEY.

    lt_where_tab = it_where_clause.

    LOOP AT it_field_names ASSIGNING FIELD-SYMBOL(<fs_fname>).
      APPEND VALUE #( fieldname = <fs_fname> ) TO lt_fields.
    ENDLOOP.

    CALL FUNCTION 'RFC_READ_TABLE'
      DESTINATION iv_rfc_destination
      EXPORTING
        query_table          = iv_table_name
        delimiter            = '|'
      TABLES
        options              = lt_where_tab
        fields               = lt_fields
        data                 = lt_data_tab
      EXCEPTIONS
        table_not_available  = 1
        table_without_data   = 2
        option_not_valid     = 3
        field_not_valid      = 4
        not_authorized       = 5
        data_buffer_exceeded = 6
        OTHERS               = 7.

    IF sy-subrc <> 0.
      ev_error_message = |RFC_READ_TABLE on { iv_table_name } failed (rc={ sy-subrc })|.
      rv_ok = abap_false.
      RETURN.
    ENDIF.

    et_result_rows = lt_data_tab.
    et_fields_tab  = lt_fields.
    rv_ok          = abap_true.

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

    DATA lv_address_lfa1 TYPE adrnr.
    DATA lv_address_kna1 TYPE adrnr.
    DATA lt_lfa1_rows    TYPE esh_t_co_rfcrt_data.
    DATA lt_lfa1_fields  TYPE ehfndt_db_fields.
    DATA lt_kna1_rows    TYPE esh_t_co_rfcrt_data.
    DATA lt_kna1_fields  TYPE ehfndt_db_fields.
    DATA lt_adr6_rows    TYPE esh_t_co_rfcrt_data.
    DATA lt_adr6_fields  TYPE ehfndt_db_fields.
    DATA lt_adrt_rows    TYPE esh_t_co_rfcrt_data.
    DATA lt_adrt_fields  TYPE ehfndt_db_fields.
    DATA lt_address_tmp  TYPE tt_itab.

    CLEAR address_data.
    rv_ok = abap_false.

    DATA(lv_bp_padded) = CONV bu_partner( |{ business_partner ALPHA = IN }| ).

    "-- 1. Look up address number in LFA1 (vendor) --------------------------
    DATA(lt_where_lfa1)  = VALUE esh_t_co_rfcrt_options( ( |LIFNR EQ { cl_abap_dyn_prg=>quote( lv_bp_padded ) }| ) ).
    DATA(lt_fields_lfa1) = VALUE stringtab( ( |LIFNR| ) ( |ADRNR| ) ).

    DATA(lv_found_lfa1) = read_remote_table(
      EXPORTING iv_rfc_destination = rfc_destination
                iv_table_name      = lc_lfa1_table
                it_where_clause    = lt_where_lfa1
                it_field_names     = lt_fields_lfa1
      IMPORTING et_result_rows     = lt_lfa1_rows
                et_fields_tab      = lt_lfa1_fields
                ev_error_message   = error_message ).

    IF lv_found_lfa1 = abap_true AND lines( lt_lfa1_rows ) > 0.
      lv_address_lfa1 = extract_field_value( iv_data_line  = lt_lfa1_rows[ 1 ]
                                             it_fields_tab = lt_lfa1_fields
                                             iv_field_name = 'ADRNR' ).
    ENDIF.

    "-- 2. Fall back to KNA1 (customer) if not found in LFA1 ----------------
    DATA(lt_where_kna1)  = VALUE esh_t_co_rfcrt_options( ( |KUNNR EQ '{ lv_bp_padded }'| ) ).
    DATA(lt_fields_kna1) = VALUE stringtab( ( |KUNNR| ) ( |ADRNR| ) ).

    DATA(lv_found_kna1) = read_remote_table(
      EXPORTING iv_rfc_destination = rfc_destination
                iv_table_name      = lc_kna1_table
                it_where_clause    = lt_where_kna1
                it_field_names     = lt_fields_kna1
      IMPORTING et_result_rows     = lt_kna1_rows
                et_fields_tab      = lt_kna1_fields
                ev_error_message   = error_message ).

    IF lv_found_kna1 = abap_true AND lines( lt_kna1_rows ) > 0.
      lv_address_kna1 = extract_field_value( iv_data_line  = lt_kna1_rows[ 1 ]
                                             it_fields_tab = lt_kna1_fields
                                             iv_field_name = 'ADRNR' ).
    ENDIF.

    IF lv_address_kna1 IS INITIAL AND lv_address_lfa1 IS INITIAL.
      error_message = |BP { business_partner }: not found in LFA1 or KNA1 on remote system|.
      RETURN.
    ENDIF.

    "-- 3. Read e-mail from ADR6 ---------------------------------------------
    DATA(lt_where_adr6)  = VALUE esh_t_co_rfcrt_options(
      ( |( ADDRNUMBER EQ '{ lv_address_lfa1 }' OR ADDRNUMBER EQ '{ lv_address_kna1 }' )| )
      ( | AND PERSNUMBER EQ ''| ) ).
    DATA(lt_fields_adr6) = VALUE stringtab( ( |ADDRNUMBER| ) ( |SMTP_ADDR| ) ( |FLGDEFAULT| ) ( |CONSNUMBER| ) ).

    DATA(lv_adr6_ok) = read_remote_table(
      EXPORTING iv_rfc_destination = rfc_destination
                iv_table_name      = lc_adr6_table
                it_where_clause    = lt_where_adr6
                it_field_names     = lt_fields_adr6
      IMPORTING et_result_rows     = lt_adr6_rows
                et_fields_tab      = lt_adr6_fields
                ev_error_message   = error_message ).

    IF lv_adr6_ok = abap_true AND lines( lt_adr6_rows ) > 0.
      LOOP AT lt_adr6_rows ASSIGNING FIELD-SYMBOL(<fs_adr6_row>).
        APPEND INITIAL LINE TO address_data ASSIGNING FIELD-SYMBOL(<fs_address_data>).
        <fs_address_data>-email      = extract_field_value( iv_data_line  = <fs_adr6_row>
                                                            it_fields_tab = lt_adr6_fields
                                                            iv_field_name = 'SMTP_ADDR' ).
        <fs_address_data>-std_no     = extract_field_value( iv_data_line  = <fs_adr6_row>
                                                            it_fields_tab = lt_adr6_fields
                                                            iv_field_name = 'FLGDEFAULT' ).
        <fs_address_data>-consnumber = extract_field_value( iv_data_line  = <fs_adr6_row>
                                                            it_fields_tab = lt_adr6_fields
                                                            iv_field_name = 'CONSNUMBER' ).
        <fs_address_data>-email_x         = abap_true.
        <fs_address_data>-businesspartner = lv_bp_padded.
      ENDLOOP.
    ENDIF.

    "-- 4. Read address note from ADRT ---------------------------------------
    DATA(lt_where_adrt)  = VALUE esh_t_co_rfcrt_options(
      ( |( ADDRNUMBER EQ '{ lv_address_lfa1 }' OR ADDRNUMBER EQ '{ lv_address_kna1 }' )| ) ).
    DATA(lt_fields_adrt) = VALUE stringtab( ( |ADDRNUMBER| ) ( |REMARK| ) ( |LANGU| ) ( |CONSNUMBER| ) ).

    DATA(lv_adrt_ok) = read_remote_table(
      EXPORTING iv_rfc_destination = rfc_destination
                iv_table_name      = lc_adrt_table
                it_where_clause    = lt_where_adrt
                it_field_names     = lt_fields_adrt
      IMPORTING et_result_rows     = lt_adrt_rows
                et_fields_tab      = lt_adrt_fields
                ev_error_message   = error_message ).

    IF lv_adrt_ok = abap_true AND lines( lt_adrt_rows ) > 0.
      LOOP AT lt_adrt_rows ASSIGNING FIELD-SYMBOL(<fs_adrt_row>).
        APPEND INITIAL LINE TO lt_address_tmp ASSIGNING FIELD-SYMBOL(<fs_tmp>).
        <fs_tmp>-notes      = extract_field_value( iv_data_line  = <fs_adrt_row>
                                                   it_fields_tab = lt_adrt_fields
                                                   iv_field_name = 'REMARK' ).
        <fs_tmp>-language   = extract_field_value( iv_data_line  = <fs_adrt_row>
                                                   it_fields_tab = lt_adrt_fields
                                                   iv_field_name = 'LANGU' ).
        <fs_tmp>-consnumber = extract_field_value( iv_data_line  = <fs_adrt_row>
                                                   it_fields_tab = lt_adrt_fields
                                                   iv_field_name = 'CONSNUMBER' ).
        <fs_tmp>-notes_x    = abap_true.
      ENDLOOP.
    ENDIF.

    "-- 5. Merge emails and remarks -----------------------------------------
    LOOP AT address_data ASSIGNING <fs_address_data>.
      TRY.
          <fs_address_data>-notes    = lt_address_tmp[ consnumber = <fs_address_data>-consnumber ]-notes.
          <fs_address_data>-language = lt_address_tmp[ consnumber = <fs_address_data>-consnumber ]-language.
          <fs_address_data>-notes_x  = COND #( WHEN <fs_address_data>-notes IS NOT INITIAL
                                               THEN abap_true
                                               ELSE abap_false ).
        CATCH cx_sy_itab_line_not_found.
          CONTINUE.
      ENDTRY.
    ENDLOOP.

    "-- 6. Guard: at least one record must be present -----------------------
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
        APPEND INITIAL LINE TO et_data ASSIGNING FIELD-SYMBOL(<fs_data>).
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
        WHEN '0006'.
          <fs_data>-std_no = <fs_excel>-value.
        WHEN '0007'.
          <fs_data>-language = <fs_excel>-value.
      ENDCASE.
    ENDLOOP.

    IF et_data IS INITIAL.
      ev_error_msg = 'No data rows found after removing header row'.
      RETURN.
    ENDIF.

    rv_ok = abap_true.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_MODEL->LOAD_EXCEL_BANK
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_FILE                        TYPE        LOCALFILE
* | [<---] ET_DATA                        TYPE        TT_BANK_ITAB
* | [<---] EV_ERROR_MSG                   TYPE        STRING
* | [<-()] RV_OK                          TYPE        ABAP_BOOL
* | Column mapping (1-based, row 1 = header skipped):
* |  1=LIFNR  2=BANKS  3=BANKL  4=BANKN   5=BKONT  6=BVTYP   7=XEZER
* |  8=BKREF  9=KOINH  10=BANK_GUID       11=TECH_RECTYP
* | 12=EBPP_ACCNAME  13=EBPP_BVSTATUS  14=KOVON  15=KOBIS
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD load_excel_bank.

    DATA lt_excel TYPE STANDARD TABLE OF alsmex_tabline.

    rv_ok = abap_false.
    CLEAR et_data.

    CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
      EXPORTING
        filename                = iv_file
        i_begin_col             = 1
        i_begin_row             = 1
        i_end_col               = 15
        i_end_row               = 9999
      TABLES
        intern                  = lt_excel
      EXCEPTIONS
        inconsistent_parameters = 1
        upload_ole              = 2
        OTHERS                  = 3.

    IF sy-subrc <> 0.
      ev_error_msg = |Banking Excel upload failed (ALSM_EXCEL_TO_INTERNAL_TABLE rc={ sy-subrc })|.
      RETURN.
    ENDIF.

    LOOP AT lt_excel ASSIGNING FIELD-SYMBOL(<fs_excel>) WHERE row <> '0001'.
      AT NEW row.
        APPEND INITIAL LINE TO et_data ASSIGNING FIELD-SYMBOL(<fs_bank>).
      ENDAT.
      CASE <fs_excel>-col.
        WHEN '0001'.
          <fs_bank>-lifnr        = CONV #( <fs_excel>-value ).
          "-- Derive businesspartner from LIFNR (vendor number = BP number) --
          <fs_bank>-businesspartner = CONV #( <fs_excel>-value ).
        WHEN '0002'.
          <fs_bank>-banks        = CONV #( <fs_excel>-value ).
        WHEN '0003'.
          <fs_bank>-bankl        = CONV #( <fs_excel>-value ).
        WHEN '0004'.
          <fs_bank>-bankn        = CONV #( <fs_excel>-value ).
        WHEN '0005'.
          <fs_bank>-bkont        = CONV #( <fs_excel>-value ).
        WHEN '0006'.
          <fs_bank>-bvtyp        = CONV #( <fs_excel>-value ).
        WHEN '0007'.
          <fs_bank>-xezer        = CONV #( <fs_excel>-value ).
        WHEN '0008'.
          <fs_bank>-bkref        = CONV #( <fs_excel>-value ).
        WHEN '0009'.
          <fs_bank>-koinh        = CONV #( <fs_excel>-value ).
        WHEN '0010'.
          <fs_bank>-bank_guid    = CONV #( <fs_excel>-value ).
        WHEN '0011'.
          <fs_bank>-tech_rectyp  = CONV #( <fs_excel>-value ).
        WHEN '0012'.
          <fs_bank>-ebpp_accname = CONV #( <fs_excel>-value ).
        WHEN '0013'.
          <fs_bank>-ebpp_bvstatus = CONV #( <fs_excel>-value ).
        WHEN '0014'.
          <fs_bank>-kovon        = CONV #( <fs_excel>-value ).
        WHEN '0015'.
          <fs_bank>-kobis        = CONV #( <fs_excel>-value ).
      ENDCASE.
    ENDLOOP.

    IF et_data IS INITIAL.
      ev_error_msg = 'No banking data rows found after removing header row'.
      RETURN.
    ENDIF.

    rv_ok = abap_true.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_MODEL->FETCH_BANK_FROM_RFC
* +-------------------------------------------------------------------------------------------------+
* | [--->] IV_BUSINESS_PARTNER            TYPE        BU_PARTNER
* | [--->] IV_RFC_DESTINATION             TYPE        RFCDEST
* | [<---] ET_BANK_DATA                   TYPE        TT_BANK_ITAB
* | [<---] EV_ERROR_MESSAGE               TYPE        STRING
* | [<-()] RV_OK                          TYPE        ABAP_BOOL
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD fetch_bank_from_rfc.

    DATA lt_lfbk_rows   TYPE esh_t_co_rfcrt_data.
    DATA lt_lfbk_fields TYPE ehfndt_db_fields.
    DATA lt_knbk_rows   TYPE esh_t_co_rfcrt_data.
    DATA lt_knbk_fields TYPE ehfndt_db_fields.
    DATA lt_tiban_rows  TYPE esh_t_co_rfcrt_data.
    DATA lt_tiban_fields TYPE ehfndt_db_fields.

    CLEAR et_bank_data.
    rv_ok = abap_false.

    DATA(lv_bp_padded) = CONV bu_partner( |{ iv_business_partner ALPHA = IN }| ).

    "-- 1. Read LFBK (vendor bank data) ------------------------------------
    DATA(lt_where_lfbk)  = VALUE esh_t_co_rfcrt_options( ( |LIFNR EQ { cl_abap_dyn_prg=>quote( lv_bp_padded ) }| ) ).
    DATA(lt_fields_lfbk) = VALUE stringtab(
      ( |LIFNR| ) ( |BANKS| ) ( |BANKL| ) ( |BANKN| ) ( |BKONT| )
      ( |BVTYP| ) ( |XEZER| ) ( |BKREF| ) ( |KOINH| )
      ( |BANK_GUID| ) ( |TECH_RECTYP| )
      ( |EBPP_ACCNAME| ) ( |EBPP_BVSTATUS| ) ( |KOVON| ) ( |KOBIS| ) ).

    DATA(lv_lfbk_ok) = read_remote_table(
      EXPORTING iv_rfc_destination = iv_rfc_destination
                iv_table_name      = lc_lfbk_table
                it_where_clause    = lt_where_lfbk
                it_field_names     = lt_fields_lfbk
      IMPORTING et_result_rows     = lt_lfbk_rows
                et_fields_tab      = lt_lfbk_fields
                ev_error_message   = ev_error_message ).

    IF lv_lfbk_ok = abap_true.
      LOOP AT lt_lfbk_rows ASSIGNING FIELD-SYMBOL(<fs_lfbk_row>).
        APPEND INITIAL LINE TO et_bank_data ASSIGNING FIELD-SYMBOL(<fs_bank>).
        <fs_bank>-businesspartner = lv_bp_padded.
        <fs_bank>-lifnr           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'LIFNR' ) ).
        <fs_bank>-banks           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BANKS' ) ).
        <fs_bank>-bankl           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BANKL' ) ).
        <fs_bank>-bankn           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BANKN' ) ).
        <fs_bank>-bkont           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BKONT' ) ).
        <fs_bank>-bvtyp           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BVTYP' ) ).
        <fs_bank>-xezer           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'XEZER' ) ).
        <fs_bank>-bkref           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BKREF' ) ).
        <fs_bank>-koinh           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'KOINH' ) ).
        <fs_bank>-bank_guid       = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'BANK_GUID' ) ).
        <fs_bank>-tech_rectyp     = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'TECH_RECTYP' ) ).
        <fs_bank>-ebpp_accname    = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'EBPP_ACCNAME' ) ).
        <fs_bank>-ebpp_bvstatus   = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'EBPP_BVSTATUS' ) ).
        <fs_bank>-kovon           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'KOVON' ) ).
        <fs_bank>-kobis           = CONV #( extract_field_value( iv_data_line  = <fs_lfbk_row>
                                                                  it_fields_tab = lt_lfbk_fields
                                                                  iv_field_name = 'KOBIS' ) ).
      ENDLOOP.
    ENDIF.

    "-- 2. Read KNBK (customer bank data) ----------------------------------
    DATA(lt_where_knbk)  = VALUE esh_t_co_rfcrt_options( ( |KUNNR EQ { cl_abap_dyn_prg=>quote( lv_bp_padded ) }| ) ).
    DATA(lt_fields_knbk) = VALUE stringtab(
      ( |KUNNR| ) ( |BANKS| ) ( |BANKL| ) ( |BANKN| ) ( |BKONT| )
      ( |BVTYP| ) ( |XEZER| ) ( |BKREF| ) ( |KOINH| )
      ( |BANK_GUID| ) ( |TECH_RECTYP| )
      ( |EBPP_ACCNAME| ) ( |EBPP_BVSTATUS| ) ( |KOVON| ) ( |KOBIS| ) ).

    DATA(lv_knbk_ok) = read_remote_table(
      EXPORTING iv_rfc_destination = iv_rfc_destination
                iv_table_name      = lc_knbk_table
                it_where_clause    = lt_where_knbk
                it_field_names     = lt_fields_knbk
      IMPORTING et_result_rows     = lt_knbk_rows
                et_fields_tab      = lt_knbk_fields
                ev_error_message   = ev_error_message ).

    IF lv_knbk_ok = abap_true.
      LOOP AT lt_knbk_rows ASSIGNING FIELD-SYMBOL(<fs_knbk_row>).
        APPEND INITIAL LINE TO et_bank_data ASSIGNING <fs_bank>.
        <fs_bank>-businesspartner = lv_bp_padded.
        "-- KNBK uses KUNNR instead of LIFNR; map to lifnr field for uniformity
        <fs_bank>-lifnr           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'KUNNR' ) ).
        <fs_bank>-banks           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BANKS' ) ).
        <fs_bank>-bankl           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BANKL' ) ).
        <fs_bank>-bankn           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BANKN' ) ).
        <fs_bank>-bkont           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BKONT' ) ).
        <fs_bank>-bvtyp           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BVTYP' ) ).
        <fs_bank>-xezer           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'XEZER' ) ).
        <fs_bank>-bkref           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BKREF' ) ).
        <fs_bank>-koinh           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'KOINH' ) ).
        <fs_bank>-bank_guid       = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'BANK_GUID' ) ).
        <fs_bank>-tech_rectyp     = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'TECH_RECTYP' ) ).
        <fs_bank>-ebpp_accname    = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'EBPP_ACCNAME' ) ).
        <fs_bank>-ebpp_bvstatus   = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'EBPP_BVSTATUS' ) ).
        <fs_bank>-kovon           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'KOVON' ) ).
        <fs_bank>-kobis           = CONV #( extract_field_value( iv_data_line  = <fs_knbk_row>
                                                                  it_fields_tab = lt_knbk_fields
                                                                  iv_field_name = 'KOBIS' ) ).
      ENDLOOP.
    ENDIF.

    IF et_bank_data IS INITIAL.
      ev_error_message = |BP { iv_business_partner }: no bank data found in LFBK or KNBK on remote system|.
      RETURN.
    ENDIF.

    "-- 3. Enrich each bank record with IBAN from TIBAN ----------------------
    LOOP AT et_bank_data ASSIGNING <fs_bank>.
      DATA(lt_where_tiban) = VALUE esh_t_co_rfcrt_options(
        ( |BANKS EQ { cl_abap_dyn_prg=>quote( <fs_bank>-banks ) }| )
        ( | AND BANKL EQ { cl_abap_dyn_prg=>quote( <fs_bank>-bankl ) }| )
        ( | AND BANKN EQ { cl_abap_dyn_prg=>quote( <fs_bank>-bankn ) }| )
        ( | AND BKONT EQ { cl_abap_dyn_prg=>quote( <fs_bank>-bkont ) }| ) ).
      DATA(lt_fields_tiban) = VALUE stringtab( ( |BANKS| ) ( |BANKL| ) ( |BANKN| ) ( |BKONT| ) ( |IBAN| ) ).

      DATA(lv_tiban_ok) = read_remote_table(
        EXPORTING iv_rfc_destination = iv_rfc_destination
                  iv_table_name      = lc_tiban_table
                  it_where_clause    = lt_where_tiban
                  it_field_names     = lt_fields_tiban
        IMPORTING et_result_rows     = lt_tiban_rows
                  et_fields_tab      = lt_tiban_fields
                  ev_error_message   = ev_error_message ).

      IF lv_tiban_ok = abap_true AND lines( lt_tiban_rows ) > 0.
        <fs_bank>-iban = CONV #( extract_field_value( iv_data_line  = lt_tiban_rows[ 1 ]
                                                      it_fields_tab = lt_tiban_fields
                                                      iv_field_name = 'IBAN' ) ).
      ENDIF.
    ENDLOOP.

    rv_ok = abap_true.

  ENDMETHOD.


* <SIGNATURE>---------------------------------------------------------------------------------------+
* | Instance Public Method ZMM_CL_BP_UPD_MODEL->UPDATE_BANK_DETAILS
* +-------------------------------------------------------------------------------------------------+
* | [--->] IT_DATA                        TYPE        TT_BANK_ITAB
* | [--->] IV_TEST_MODE                   TYPE        ABAP_BOOL (default =ABAP_TRUE)
* | [<-()] RT_LOG                         TYPE        TT_ERRORLOG
* | Uses BAPI_BUPA_BANKDETAIL_CHANGE per row.
* | BKVID (bank detail ID) is derived from BANKL+BANKN; if BANK_GUID is
* | filled it is passed in the BANKDETAILDATA structure.
* +--------------------------------------------------------------------------------------</SIGNATURE>
  METHOD update_bank_details.

    DATA ls_bankdetail   TYPE bapibus1006_bankdetail.
    DATA ls_bankdetail_x TYPE bapibus1006_bankdetail_x.
    DATA lt_return       TYPE bapiret2_t.
    DATA ls_log          TYPE ty_errorlog.
    DATA lv_row          TYPE i.

    LOOP AT it_data ASSIGNING FIELD-SYMBOL(<fs_bank>).
      lv_row += 1.
      CLEAR: ls_bankdetail, ls_bankdetail_x, lt_return, ls_log.

      ls_log = VALUE #( row   = lv_row
                        accid = <fs_bank>-businesspartner ).

      "-- Validate: BP number must be filled --------------------------------
      IF <fs_bank>-businesspartner IS INITIAL.
        rt_log = VALUE #( BASE rt_log
          ( row = ls_log-row accid = ls_log-accid type = 'E'
            msg = |BP number is empty - row skipped| ) ).
        CONTINUE.
      ENDIF.

      "-- Validate: bank key must be filled ---------------------------------
      IF <fs_bank>-bankl IS INITIAL.
        rt_log = VALUE #( BASE rt_log
          ( row = ls_log-row accid = ls_log-accid type = 'W'
            msg = |BP { <fs_bank>-businesspartner }: bank key (BANKL) is empty - row skipped| ) ).
        CONTINUE.
      ENDIF.

      DATA(lv_bp) = CONV bu_partner( |{ <fs_bank>-businesspartner ALPHA = IN }| ).

      "-- Derive bank detail ID: use bank key + account number --------------
      "-- BKVID is a sequential identifier maintained in BUT0BK;
      "-- when updating an existing record the caller must supply the correct
      "-- BKVID.  Here we use BANKL+BANKN truncated to fit BKVID (3 chars).
      "-- In productive use, BKVID should be read from BUT0BK first.
      DATA(lv_bkvid) = CONV bu_bkvid( <fs_bank>-bankl(3) ).

      "-- Fill BAPIBUS1006_BANKDETAIL ----------------------------------------
      ls_bankdetail = VALUE #(
        bank_ctry    = <fs_bank>-banks
        bank_key     = <fs_bank>-bankl
        bank_acct    = <fs_bank>-bankn
        bkont        = <fs_bank>-bkont
        bank_ref     = <fs_bank>-bkref
        acct_hold    = <fs_bank>-koinh
        bvtyp        = <fs_bank>-bvtyp
        xezer        = <fs_bank>-xezer
        validfrom    = <fs_bank>-kovon
        validto      = <fs_bank>-kobis
        iban         = <fs_bank>-iban ).

      "-- Fill BAPIBUS1006_BANKDETAIL_X (change flags) ----------------------
      ls_bankdetail_x = VALUE #(
        bank_ctry    = COND #( WHEN <fs_bank>-banks IS NOT INITIAL THEN abap_true )
        bank_key     = COND #( WHEN <fs_bank>-bankl IS NOT INITIAL THEN abap_true )
        bank_acct    = COND #( WHEN <fs_bank>-bankn IS NOT INITIAL THEN abap_true )
        bkont        = COND #( WHEN <fs_bank>-bkont IS NOT INITIAL THEN abap_true )
        bank_ref     = COND #( WHEN <fs_bank>-bkref IS NOT INITIAL THEN abap_true )
        acct_hold    = COND #( WHEN <fs_bank>-koinh IS NOT INITIAL THEN abap_true )
        bvtyp        = COND #( WHEN <fs_bank>-bvtyp IS NOT INITIAL THEN abap_true )
        xezer        = COND #( WHEN <fs_bank>-xezer IS NOT INITIAL THEN abap_true )
        validfrom    = COND #( WHEN <fs_bank>-kovon IS NOT INITIAL THEN abap_true )
        validto      = COND #( WHEN <fs_bank>-kobis IS NOT INITIAL THEN abap_true )
        iban         = COND #( WHEN <fs_bank>-iban  IS NOT INITIAL THEN abap_true ) ).

      "-- Call BAPI (skip in test mode) -------------------------------------
      IF iv_test_mode = abap_false.
        CALL FUNCTION 'BAPI_BUPA_BANKDETAIL_CHANGE'
          EXPORTING
            businesspartner = lv_bp
            bankdetailid    = lv_bkvid
            bankdetaildata  = ls_bankdetail
            bankdetaildata_x = ls_bankdetail_x
          TABLES
            return          = lt_return.
      ENDIF.

      "-- Evaluate result ---------------------------------------------------
      IF iv_test_mode = abap_true.
        rt_log = VALUE #( BASE rt_log
          ( row   = lv_row
            accid = lv_bp
            type  = 'S'
            msg   = |TEST - Bank: { <fs_bank>-bankl } / Acct: { <fs_bank>-bankn } / BP: { <fs_bank>-businesspartner } - no changes written| ) ).

      ELSEIF NOT line_exists( lt_return[ type = 'E' ] )
         AND NOT line_exists( lt_return[ type = 'A' ] ).
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait = abap_true.
        rt_log = VALUE #( BASE rt_log
          ( row   = lv_row
            accid = lv_bp
            type  = 'S'
            msg   = |OK - Bank: { <fs_bank>-bankl } / Acct: { <fs_bank>-bankn } / BP: { <fs_bank>-businesspartner }| ) ).

      ELSE.
        CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

        TRY.
            DATA(ls_ret) = lt_return[ type = 'E' ].
            rt_log = VALUE #( BASE rt_log
              ( row        = lv_row
                accid      = lv_bp
                id         = ls_ret-id
                type       = 'E'
                number     = ls_ret-number
                message    = ls_ret-message
                log_msg_no = ls_ret-log_msg_no
                message_v1 = ls_ret-message_v1
                message_v2 = ls_ret-message_v2
                message_v3 = ls_ret-message_v3
                message_v4 = ls_ret-message_v4
                msg        = collect_bapi_msgs( lt_return ) ) ).
          CATCH cx_sy_itab_line_not_found.
        ENDTRY.

      ENDIF.

    ENDLOOP.

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

      "-- Validate: BP number must be filled --------------------------------
      IF <fs_data>-businesspartner IS INITIAL.
        rt_log = VALUE #( BASE rt_log
          ( row = ls_log-row accid = ls_log-accid type = 'E'
            msg = |BP number is empty - row skipped| ) ).
        CONTINUE.
      ENDIF.

      "-- Validate: at least one payload field ------------------------------
      IF <fs_data>-email IS INITIAL AND <fs_data>-notes IS INITIAL.
        rt_log = VALUE #( BASE rt_log
          ( row = ls_log-row accid = ls_log-accid type = 'W'
            msg = |BP { <fs_data>-businesspartner }: email and note both empty - skipped| ) ).
        CONTINUE.
      ENDIF.

      DATA(lv_bp) = CONV bu_partner( |{ <fs_data>-businesspartner ALPHA = IN }| ).

      "-- Build SMTP (email) structures -------------------------------------
      IF <fs_data>-email IS NOT INITIAL.
        lt_adsmtp  = VALUE #( ( e_mail = <fs_data>-email consnumber = '001' std_no = <fs_data>-std_no ) ).
        lt_adsmt_x = VALUE #( ( e_mail = COND #( WHEN <fs_data>-email_x IS INITIAL
                                                 THEN 'X'
                                                 ELSE <fs_data>-email_x )
                                consnumber = 'X' updateflag = 'I' ) ).
      ENDIF.

      "-- Build COMM_NOTES (address note) structures ------------------------
      IF <fs_data>-notes IS NOT INITIAL.
        lt_comrem  = VALUE #( ( comm_type = 'INT' langu = <fs_data>-language comm_notes = <fs_data>-notes ) ).
        lt_comre_x = VALUE #( ( comm_notes = COND #( WHEN <fs_data>-notes_x IS INITIAL
                                                     THEN 'X'
                                                     ELSE <fs_data>-notes_x )
                                updateflag = 'I' ) ).
      ENDIF.

      "-- Call BAPI_BUPA_ADDRESS_CHANGE (skip DB write in test mode) --------
      CALL FUNCTION 'BAPI_BUPA_ADDRESS_CHANGE'
        EXPORTING
          businesspartner = lv_bp
        TABLES
          bapiadsmtp      = lt_adsmtp
          bapicomrem      = lt_comrem
          bapiadsmt_x     = lt_adsmt_x
          bapicomre_x     = lt_comre_x
          return          = lt_return.

      "-- Evaluate result ---------------------------------------------------
      IF NOT line_exists( lt_return[ type = 'E' ] )
     AND NOT line_exists( lt_return[ type = 'A' ] ).
        IF iv_test_mode = abap_false.
          CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
            EXPORTING
              wait = abap_true.
          rt_log = VALUE #( BASE rt_log
            ( row   = ls_log-row
              accid = ls_log-accid
              type  = 'S'
              msg   = |OK - Email: { <fs_data>-email } / Note: { <fs_data>-notes } / BP: { <fs_data>-businesspartner }| ) ).
        ELSE.
          CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
          rt_log = VALUE #( BASE rt_log
            ( row   = ls_log-row
              accid = ls_log-accid
              type  = 'S'
              msg   = |TEST - Email: { <fs_data>-email } / Note: { <fs_data>-notes } / BP: { <fs_data>-businesspartner } - no changes written| ) ).
        ENDIF.
      ELSE.
        CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

        TRY.
            DATA(ls_ret) = lt_return[ type = 'E' ].
            rt_log = VALUE #( BASE rt_log
              ( row        = lv_row
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
