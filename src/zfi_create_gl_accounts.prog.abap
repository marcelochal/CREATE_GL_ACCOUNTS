*&---------------------------------------------------------------------*
*& Report ZFI_CREATE_GL_ACCOUNTS
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zfi_create_gl_accounts.

TABLES: sscrfields.
TYPE-POOLS icon.

TYPES: BEGIN OF ty_upload,
         saknr          TYPE glaccount_name-keyy-saknr, "Nº conta do Razão
         bukrs          TYPE glaccount_ccode-keyy-bukrs,
         glaccount_type TYPE glaccount_coa-data-glaccount_type,
         ktoks          TYPE glaccount_coa-data-ktoks,
         txt20          TYPE glaccount_name-data-txt20,
         txt50          TYPE glaccount_name-data-txt50,
         waers          TYPE glaccount_ccode-data-waers,
         xsalh          TYPE glaccount_ccode-data-xsalh,
         mwskz          TYPE glaccount_ccode-data-mwskz,
         xmwno          TYPE glaccount_ccode-data-xmwno,
         mitkz          TYPE glaccount_ccode-data-mitkz,
         xopvw          TYPE glaccount_ccode-data-xopvw,
         katyp          TYPE glaccount_carea-data-katyp,
         mgefl          TYPE glaccount_carea-data-mgefl,
         fstag          TYPE glaccount_ccode-data-fstag,
         xintb          TYPE glaccount_ccode-data-xintb,
         zuawa          TYPE glaccount_ccode-data-zuawa,
         bilkt          TYPE glaccount_coa-data-bilkt,
       END OF ty_upload,

       BEGIN OF ty_result,
         status TYPE tp_icon,
         msg    TYPE bapi_msg,
       END OF   ty_result.


*       BEGIN OF ty_message,
*         saknr   TYPE saknr,       " G/L Account
*         bukrs   TYPE bukrs,       " Company Code
*         type    TYPE bapi_mtype,  " Message type: S Success, E Error, W Warning, I Info, A Abort
*         message TYPE bapi_msg,    " Message Text
*       END OF ty_message.

DATA BEGIN OF es_alv.
INCLUDE TYPE ty_result.
INCLUDE TYPE ty_upload.
DATA END OF es_alv.

DATA:
      it_alv LIKE STANDARD TABLE OF es_alv.

CONSTANTS:
      cc_parameter_id TYPE memoryid VALUE 'Z_FILE_01'.

*&---------------------------------------------------------------------*
*& Class definition lcl_file
*&---------------------------------------------------------------------*
CLASS lcl_file DEFINITION FINAL.
  PUBLIC SECTION.

    CLASS-METHODS:
      set_init_filename RETURNING VALUE(r_filename) TYPE file_table-filename,
      select            RETURNING VALUE(r_filename) TYPE file_table-filename,
      upload EXCEPTIONS file_path_is_initial upload_failed,
      check_path.

    CLASS-DATA:
      get_file          TYPE filename,
      get_table         TYPE TABLE OF ty_upload,
      get_upload_path   TYPE string,
      get_download_path TYPE string.

  PROTECTED SECTION.

ENDCLASS.
*&---------------------------------------------------------------------*
*& Class definition Messages
*&---------------------------------------------------------------------*
CLASS lcl_messages DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS:
      initialize,
      store,
      show.

ENDCLASS.
*&---------------------------------------------------------------------*
*& Class definition GL_ACCOUNTS
*&---------------------------------------------------------------------*
CLASS lcl_gl_accounts DEFINITION FINAL FRIENDS lcl_messages.
  PUBLIC SECTION.
    CLASS-METHODS:
      create.

    CLASS-DATA:
      it_carea   TYPE glaccount_carea_table,
      it_ccodes  TYPE glaccount_ccode_table,
      it_coa     TYPE TABLE OF glaccount_coa,
      it_keyword TYPE glaccount_keyword_table,
      it_names   TYPE glaccount_name_table,
      it_return  TYPE TABLE OF bapiret2,
      it_upload  TYPE TABLE OF ty_upload.

ENDCLASS.

**********************************************************************
*SELECTION-SCREEN
**********************************************************************
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
PARAMETERS:
  p_ktopl TYPE ktopl  DEFAULT 'TB00',
  p_kokrs TYPE kokrs  DEFAULT 'TB00' MATCHCODE OBJECT csh_tka01,
  p_file  TYPE localfile LOWER CASE.
SELECTION-SCREEN: FUNCTION KEY 1.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
PARAMETERS p_test TYPE xtest DEFAULT abap_true.
SELECTION-SCREEN END OF BLOCK b2.


**********************************************************************
*INITIALIZATION
**********************************************************************
INITIALIZATION.
  lcl_file=>set_init_filename( ).
  lcl_messages=>initialize( ).

  sscrfields-functxt_01 = icon_export && | Baixar modelo| .

**********************************************************************
*SELECTION-SCREEN EVENT
**********************************************************************
AT SELECTION-SCREEN ON VALUE-REQUEST FOR  p_file.
*  PERFORM get_file_name USING p_file.
  p_file = lcl_file=>select( ).

AT SELECTION-SCREEN.
  CASE sscrfields-ucomm.
    WHEN'FC01'.
      PERFORM exporta_modelo.
  ENDCASE.

START-OF-SELECTION.

  CALL METHOD lcl_file=>upload
    EXCEPTIONS
      file_path_is_initial = 1
      upload_failed        = 2.

  IF sy-subrc NE 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
      DISPLAY LIKE 'E'.
    STOP.
  ENDIF.

  PERFORM alv_exibe USING p_test.

  CALL METHOD lcl_gl_accounts=>create( ).

END-OF-SELECTION.

*&---------------------------------------------------------------------*
*& Form ALV_EXIBE
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
FORM alv_exibe USING p_test TYPE xtest.

  DATA:
    it_fieldcat   TYPE slis_t_fieldcat_alv,
    wa_layout     TYPE slis_layout_alv,
    wl_fieldcat   TYPE LINE OF slis_t_fieldcat_alv,
    it_events     TYPE slis_t_event,
    lv_title      TYPE lvc_title,
    lv_gui_status TYPE slis_formname,
    lv_no_out(1)  TYPE c.

  PERFORM alv_create_fieldcatalog USING es_alv CHANGING it_fieldcat.

*Verifica o tipo de resultado a ser exibido
  IF p_test IS INITIAL.
    lv_title = 'Dados importados e resultado do processamento de teste'(002).
    lv_gui_status = 'ALV_SET_STATUS_01'.
    lv_no_out = abap_true.

  ELSE.
    lv_title = 'Resultado do processamento da carga'(003).
    lv_gui_status = space.
    lv_no_out = space.
  ENDIF.

  wa_layout-colwidth_optimize = abap_true.

  READ TABLE it_fieldcat INTO wl_fieldcat INDEX 2.
*  wl_fieldcat-no_out = lv_no_out.
*  wl_fieldcat-no_zero = abap_true.
  wl_fieldcat-outputlen = 30.
  MODIFY it_fieldcat FROM wl_fieldcat INDEX 2.

  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      i_callback_program       = syst-cprog
      i_callback_pf_status_set = lv_gui_status
      i_callback_user_command  = 'USER_COMMAND'
      i_grid_title             = lv_title
      is_layout                = wa_layout
      it_fieldcat              = it_fieldcat
      it_events                = it_events
    TABLES
      t_outtab                 = it_alv
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.
  IF sy-subrc NE 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.



ENDFORM.
*&---------------------------------------------------------------------*
*& Form  alv_create_fieldcatalog
*&---------------------------------------------------------------------*
* Cria catalogo de campos a partir de uma tabela interna
*----------------------------------------------------------------------*
*      -->PT_TABLE     Tabela Interna
*      -->PT_FIELDCAT  Catalogo de campos
*----------------------------------------------------------------------*
FORM  alv_create_fieldcatalog
       USING     p_structure  TYPE any
       CHANGING  pt_fieldcat  TYPE slis_t_fieldcat_alv.

  DATA:
    lr_tabdescr TYPE REF TO cl_abap_structdescr,
    lr_data     TYPE REF TO data,
    lt_dfies    TYPE ddfields,
    ls_dfies    TYPE dfies,
    ls_fieldcat TYPE lvc_s_fcat,
    it_fieldcat TYPE TABLE OF lvc_s_fcat.

  CLEAR:
    pt_fieldcat.

  CREATE DATA lr_data LIKE p_structure.

  lr_tabdescr ?= cl_abap_structdescr=>describe_by_data_ref( lr_data ).

  lt_dfies = cl_salv_data_descr=>read_structdescr( lr_tabdescr ).

  LOOP AT lt_dfies INTO ls_dfies.

    CLEAR ls_fieldcat.

    MOVE-CORRESPONDING ls_dfies TO ls_fieldcat.

*Campo Status definir como icone
    IF sy-tabix EQ 1.
      ls_fieldcat-icon = abap_true.
    ENDIF.
*Campo Transação definir como one clique
    IF sy-tabix EQ 2.
      ls_fieldcat-hotspot = abap_true.
    ENDIF.

* Retira zeros a esqueda do BP Partner
    IF sy-tabix EQ 7.
      ls_fieldcat-no_zero = abap_true.
    ENDIF.

*  Define os campos como centralizado
    ls_fieldcat-just = 'C'.

    APPEND ls_fieldcat TO it_fieldcat.

  ENDLOOP.

  MOVE-CORRESPONDING it_fieldcat TO pt_fieldcat.

ENDFORM.                    "alv_create_fieldcatalog

*&---------------------------------------------------------------------*
*& Class (Implementation) lcl_file
*&---------------------------------------------------------------------*
CLASS lcl_file IMPLEMENTATION.

  METHOD set_init_filename.

* Pega o parametro do caminho definido para o usuario
    GET PARAMETER ID cc_parameter_id FIELD p_file.

* Caso o parameter não estiver definido / vazio
* pega a definição da pasta de trabalho do usuario sapgui
    IF p_file IS INITIAL.

      CALL METHOD cl_gui_frontend_services=>get_upload_download_path
        CHANGING
          upload_path   = get_upload_path
          download_path = get_download_path.

*      CONCATENATE lv_dirup 'file.xlsx' INTO r_filename.
      r_filename = get_upload_path.

    ENDIF.

  ENDMETHOD.                    "get_init_filename

  METHOD select.
    DATA:
      lt_filetable TYPE filetable,
      lv_subrc     TYPE sysubrc.

    CONSTANTS cc_file_ext(26) TYPE c VALUE '*.xlsx;*.xlsm;*.xlsb;*.xls'.

    CALL METHOD cl_gui_frontend_services=>file_open_dialog
      EXPORTING
        window_title            = | Selecionar arquivo de carga do plano de contas |
        file_filter             = | Planilhas do Microsoft Excel ({ cc_file_ext })\| { cc_file_ext }\|  { TEXT-005 } (*.*)\|*.*|
        multiselection          = abap_false
      CHANGING
        file_table              = lt_filetable
        rc                      = lv_subrc
      EXCEPTIONS
        file_open_dialog_failed = 1
        cntl_error              = 2
        error_no_gui            = 3
        not_supported_by_gui    = 4
        OTHERS                  = 5.

    IF sy-subrc <> 0.
      RETURN.
    ENDIF.

    READ TABLE lt_filetable INTO r_filename INDEX 1.

    get_file = r_filename.

  ENDMETHOD.                    "select_file

  METHOD upload.
    DATA :
           i_raw_data TYPE truxs_t_text_data.

    IF p_file IS INITIAL.
      MESSAGE 'Por favor, insira o caminho do arquivo para prosseguir'(e01)
        TYPE 'E' RAISING file_path_is_initial.
    ENDIF.

    SET PARAMETER ID cc_parameter_id FIELD p_file.

    CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
      EXPORTING
*       i_field_seperator    = ''
        i_line_header        = abap_true
        i_tab_raw_data       = i_raw_data
        i_filename           = p_file
      TABLES
        i_tab_converted_data = lcl_file=>get_table
      EXCEPTIONS
        conversion_failed    = 1
        OTHERS               = 2.

    IF sy-subrc NE 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
      WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4
        RAISING upload_failed.
    ENDIF.

    MOVE-CORRESPONDING lcl_file=>get_table TO it_alv.

  ENDMETHOD.

  METHOD check_path.

    DATA :
      lv_dir      TYPE string,
      lv_file     TYPE string,
      lv_filename TYPE string.

    lv_filename = lcl_file=>get_file.

    CALL FUNCTION 'SO_SPLIT_FILE_AND_PATH'
      EXPORTING
        full_name     = lcl_file=>get_file
      IMPORTING
        stripped_name = lv_file "file name
        file_path     = lv_dir  "directory path
      EXCEPTIONS
        x_error       = 1
        OTHERS        = 2.
    IF sy-subrc NE 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.

    IF cl_gui_frontend_services=>directory_exist( directory = lv_dir ) IS INITIAL.
      MESSAGE ID 'C$' TYPE 'E' NUMBER '155' "Não foi possível abrir o file &1&3 (&2)
        WITH p_file .
    ENDIF.

    " check file existence
    IF cl_gui_frontend_services=>file_exist( file = lv_filename ) IS INITIAL.
      MESSAGE ID 'FES' TYPE 'E' NUMBER '000'. "O file não existe
    ENDIF.

*  SPLIT lv_filename AT '.' INTO TABLE DATA(itab).


  ENDMETHOD.                    "check_path

ENDCLASS.                       "cl_file

*&---------------------------------------------------------------------*
*& Class (Implementation) cl_messages
*&---------------------------------------------------------------------*
CLASS lcl_messages IMPLEMENTATION.

  METHOD initialize.

*    Inicializa message store.

    CALL FUNCTION 'MESSAGES_ACTIVE'
      EXCEPTIONS
        not_active = 1
        OTHERS     = 2.
    IF sy-subrc EQ 1.
      CALL FUNCTION 'MESSAGES_INITIALIZE'
        EXCEPTIONS
          log_not_active       = 1
          wrong_identification = 2
          OTHERS               = 3.
    ENDIF.

    IF sy-subrc NE 0.
*    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.
  ENDMETHOD.

  METHOD show.

    CALL FUNCTION 'MESSAGES_SHOW'
      EXPORTING
*       CORRECTIONS_OPTION = ' '
*       CORRECTIONS_FUNC_TEXT       = ' '
*       MSG_SELECT_FUNC    = ' '
*       MSG_SELECT_FUNC_TEXT        = ' '
*       LINE_FROM          = ' '
*       LINE_TO            = ' '
*       OBJECT             = ' '
*       SEND_IF_ONE        = ' '
        batch_list_type    = 'J'
*       SHOW_LINNO         = 'X'
*       SHOW_LINNO_TEXT    = ' '
*       SHOW_LINNO_TEXT_LEN         = '3'
        i_use_grid         = abap_true
        i_amodal_window    = abap_false
*     IMPORTING
*       CORRECTIONS_WANTED =
*       MSG_SELECTED       =
*       E_EXIT_COMMAND     =
      EXCEPTIONS
        inconsistent_range = 1
        no_messages        = 2
        OTHERS             = 3.
    IF sy-subrc <> 0.
* Implement suitable error handling here
    ENDIF.


  ENDMETHOD.

  METHOD store .

    DATA:
         wa_bapiret2 TYPE bapiret2.

    LOOP AT lcl_gl_accounts=>it_return INTO wa_bapiret2.

      CALL FUNCTION 'MESSAGE_STORE'
        EXPORTING
          arbgb                  = wa_bapiret2-id          "sy-msgid
          msgty                  = wa_bapiret2-type        "sy-msgty
          msgv1                  = wa_bapiret2-message_v1  "sy-msgv1
          msgv2                  = wa_bapiret2-message_v2  "sy-msgv2
          msgv3                  = wa_bapiret2-message_v3  "sy-msgv3
          msgv4                  = wa_bapiret2-message_v4  "sy-msgv4
          txtnr                  = wa_bapiret2-number      "sy-msgno
          zeile                  = wa_bapiret2-row
*         arbgb                  = sy-msgid
*         msgty                  = sy-msgty
*         msgv1                  = sy-msgv1
*         msgv2                  = sy-msgv2
*         msgv3                  = sy-msgv3
*         msgv4                  = sy-msgv4
*         txtnr                  = sy-msgno
        EXCEPTIONS
          message_type_not_valid = 1
          not_active             = 2
          OTHERS                 = 3.
      IF sy-subrc NE 0.
*      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
*              WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

    ENDLOOP.

  ENDMETHOD.

ENDCLASS.
*&---------------------------------------------------------------------*
*& Class (Implementation) lcl_messages
*&---------------------------------------------------------------------*
CLASS lcl_gl_accounts IMPLEMENTATION.

  METHOD create.

    DATA:
      wa_carea   TYPE LINE OF glaccount_carea_table,
      wa_ccodes  TYPE LINE OF glaccount_ccode_table,
      wa_coa     TYPE glaccount_coa,
      wa_keyword TYPE LINE OF glaccount_keyword_table,
*      wa_message TYPE ty_message,
      wa_names   TYPE LINE OF glaccount_name_table,
      wa_return  TYPE bapiret2,
      wa_upload  TYPE ty_upload.

    LOOP AT lcl_file=>get_table INTO wa_upload.

    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = wa_upload-saknr
      IMPORTING
        output = wa_upload-saknr.

      wa_coa-keyy-saknr = wa_upload-saknr.
      wa_coa-keyy-ktopl = p_ktopl.
      wa_coa-data-ktoks = wa_upload-ktoks.
      wa_coa-data-gvtyp = 'X'.  "Tipo de conta de resultados
      wa_coa-data-glaccount_type = wa_upload-glaccount_type.
      wa_coa-info-erdat = sy-datum.
      wa_coa-info-ernam = sy-uname.
*    wa_coa-action = 'I'. "I  Inserir, U  Modificar, D  Eliminar
      wa_coa-action = 'U'. "I Inserir, U  Modificar, D  Eliminar
      APPEND wa_coa TO it_coa.

      wa_names-keyy-saknr = wa_upload-saknr.
      wa_names-keyy-ktopl = p_ktopl.
      wa_names-keyy-spras = syst-langu.
      wa_names-data-txt20 = wa_upload-txt20.
      wa_names-data-txt50 = wa_upload-txt50.
*    wa_names-action = 'I'.
      wa_names-action = 'U'.
      APPEND wa_names TO it_names.

      wa_ccodes-keyy-saknr = wa_upload-saknr.
      wa_ccodes-keyy-bukrs = wa_upload-bukrs.
      wa_ccodes-data-waers = wa_upload-waers.
      wa_ccodes-data-xsalh = wa_upload-xsalh.
      wa_ccodes-data-mwskz = wa_upload-mwskz.
      wa_ccodes-data-xmwno = wa_upload-xmwno.
      wa_ccodes-data-mitkz = wa_upload-mitkz.
      wa_ccodes-data-xopvw = wa_upload-xopvw.
      wa_ccodes-data-fstag = wa_upload-fstag.     " Field Status Group
      wa_ccodes-data-xintb = wa_upload-xintb.     " Indicator: Is Account only Posted to Automatically?
      wa_ccodes-data-zuawa = wa_upload-zuawa.     " Sort key
      wa_ccodes-action = 'I'.
      APPEND wa_ccodes TO it_ccodes.

      IF wa_upload-glaccount_type NE 'X' AND wa_upload-glaccount_type NE 'N'.
        wa_carea-keyy-saknr = wa_upload-saknr.
        wa_carea-keyy-kokrs = p_ktopl.
        wa_carea-data-mgefl = wa_upload-mgefl.

        "cost element restriction for gl account type other than P and S RR 22062017"
        IF wa_upload-glaccount_type = 'P' ."OR gwa_upload-glaccount_type = 'S'.
          wa_carea-data-katyp = wa_upload-katyp.
        ENDIF.
        "cost element restriction for gl account type other than P and S RR 22062017"
        wa_carea-keyy-saknr = wa_upload-saknr.
        wa_carea-keyy-kokrs = p_kokrs. "Área de contabilidade de custos
        wa_carea-fromto-datab = '20010101'.            " sy-datum.
        wa_carea-fromto-datbi = '99991231'.            " Valid To Date
        wa_carea-action = 'I'.
        APPEND wa_carea TO it_carea.

        wa_coa-keyy-ktopl = p_ktopl.
        wa_coa-keyy-saknr = wa_upload-saknr.
        wa_coa-data-bilkt = wa_upload-bilkt.
        APPEND wa_coa TO it_coa.
      ENDIF.

      CALL FUNCTION 'GL_ACCT_MASTER_SAVE'
        EXPORTING
          testmode           = p_test
          no_save_at_warning = abap_true
*         NO_AUTHORITY_CHECK =
*         STORE_DATA_ONLY    =
        TABLES
          account_names      = it_names
*         ACCOUNT_KEYWORDS   =
          account_ccodes     = it_ccodes
          account_careas     = it_carea
          return             = it_return
        CHANGING
          account_coa        = wa_coa.


      READ TABLE it_return INTO wa_return WITH KEY type = 'E'.
      IF sy-subrc NE 0 AND p_test IS INITIAL.
        CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
          EXPORTING
            wait = 'X'.
      ENDIF.

      IF it_return IS NOT INITIAL.
        lcl_messages=>store( ).
      ELSE.
*      wa_message-saknr = wa_upload-saknr.
*      wa_message-bukrs = wa_upload-bukrs.
*      wa_message-type = 'S'.
*      wa_message-message = 'Created'.
*      APPEND wa_message TO it_message.
*      CLEAR: wa_message.
      ENDIF.

      REFRESH: it_names, it_ccodes, it_carea, it_return, it_coa.
      CLEAR: wa_names, wa_ccodes, wa_carea, wa_return, wa_coa.
    ENDLOOP.
  ENDMETHOD.

ENDCLASS.

FORM exporta_modelo.

  DATA:
    lo_table  TYPE REF TO cl_salv_table,
    lx_xml    TYPE xstring,
    vl_return TYPE c,
    wa_sval   TYPE sval,
    lt_sval   TYPE TABLE OF sval,
    it_upload TYPE TABLE OF ty_upload,
    wa_upload TYPE ty_upload,
    vl_saknr  TYPE saknr,
    vl_ktopl  TYPE ktopl,
    vl_bukrs  TYPE bukrs,
    wa_ska1   TYPE ska1,
    wa_skat   TYPE skat,
    wa_skb1   TYPE skb1.


  CONSTANTS:
    c_default_extension TYPE string VALUE 'xlsx',
    c_default_file_name TYPE string VALUE 'modelo.xlsx',
    c_default_mask      TYPE string VALUE 'Excel (*.xlsx)|*.xlsx' ##NO_TEXT.

  MOVE:
    'SKA1'  TO wa_sval-tabname,
    'SAKNR' TO wa_sval-fieldname.
  APPEND wa_sval TO lt_sval.

  MOVE:
    'SKA1'  TO wa_sval-tabname,
    'KTOPL' TO wa_sval-fieldname,
    p_ktopl TO wa_sval-value.
  APPEND wa_sval TO lt_sval.

  CLEAR wa_sval.

  MOVE:
    'SKB1'  TO wa_sval-tabname,
    'BUKRS' TO wa_sval-fieldname.
  APPEND wa_sval TO lt_sval.

  CALL FUNCTION 'POPUP_GET_VALUES'
    EXPORTING
*     NO_VALUE_CHECK  = ' '
      popup_title     = 'Baixar modelo preenchido'
    IMPORTING
      returncode      = vl_return
    TABLES
      fields          = lt_sval
    EXCEPTIONS
      error_in_fields = 1
      OTHERS          = 2.

  CHECK vl_return IS INITIAL.

  READ TABLE lt_sval INTO wa_sval WITH KEY  fieldname = 'SAKNR'.
  vl_saknr = wa_sval-value.
  READ TABLE lt_sval INTO wa_sval WITH KEY  fieldname = 'KTOPL'.
  vl_ktopl = wa_sval-value.
  READ TABLE lt_sval INTO wa_sval WITH KEY  fieldname = 'BUKRS'.
  vl_bukrs = wa_sval-value.

  CALL FUNCTION 'READ_SKA1'
    EXPORTING
      xktopl         = vl_ktopl
      xsaknr         = vl_saknr
    IMPORTING
      xska1          = wa_ska1
      xskat          = wa_skat
    EXCEPTIONS
      key_incomplete = 1
      not_authorized = 2
      not_found      = 3
      OTHERS         = 4.

  CALL FUNCTION 'READ_SKB1'
    EXPORTING
      xbukrs         = vl_bukrs
      xsaknr         = vl_saknr
    IMPORTING
      xskb1          = wa_skb1
    EXCEPTIONS
      key_incomplete = 1
      not_authorized = 2
      not_found      = 3
      OTHERS         = 4.

  MOVE-CORRESPONDING wa_ska1 TO wa_upload.
  MOVE-CORRESPONDING wa_skat TO wa_upload.
  MOVE-CORRESPONDING wa_skb1 TO wa_upload.

  APPEND wa_upload TO it_upload.

  TRY .
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_table CHANGING t_table = it_upload ).
    CATCH cx_root.

  ENDTRY.

  lx_xml = lo_table->to_xml( xml_type = '10' ). "XLSX

  CALL FUNCTION 'XML_EXPORT_DIALOG'
    EXPORTING
      i_xml                      = lx_xml
      i_default_extension        = c_default_extension
      i_initial_directory        = lcl_file=>get_download_path
      i_default_file_name        = c_default_file_name
      i_mask                     = c_default_mask
*     I_APPLICATION              =
    EXCEPTIONS
      application_not_executable = 1
      OTHERS                     = 2.

ENDFORM.
