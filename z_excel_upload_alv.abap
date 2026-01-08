*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Convert Hex to String, Display in ALV
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

*----------------------------------------------------------------------*
* Type Definitions
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_excel_raw,
         col_a TYPE string,      " Field 1
         col_b TYPE string,      " Field 2 (Hex String)
         col_c TYPE string,      " Field 3
       END OF ty_excel_raw.

* Output Structure
TYPES: BEGIN OF ty_final,
         excel_id    TYPE string,
         xml_content TYPE string,      " Converted XML Content
         col_c       TYPE string,
         error_msg   TYPE string,
       END OF ty_final.

*----------------------------------------------------------------------*
* Data Declarations
*----------------------------------------------------------------------*
DATA: gt_excel_raw TYPE TABLE OF ty_excel_raw,
      gt_final     TYPE TABLE OF ty_final,
      gs_final     TYPE ty_final,
      gv_file      TYPE rlgrap-filename.

*----------------------------------------------------------------------*
* Selection Screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_file TYPE localfile OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
* Initialization
*----------------------------------------------------------------------*
INITIALIZATION.
  TEXT-001 = 'File Selection'.

*----------------------------------------------------------------------*
* At Selection Screen
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f_file_open.

*----------------------------------------------------------------------*
* Start of Selection
*----------------------------------------------------------------------*
START-OF-SELECTION.
  PERFORM f_upload_excel.

  IF gt_excel_raw IS INITIAL.
    MESSAGE 'No data found in Excel file.' TYPE 'S' DISPLAY LIKE 'E'.
  ELSE.
    PERFORM f_process_data.
    PERFORM f_display_alv.
  ENDIF.

*&---------------------------------------------------------------------*
*& Form f_file_open
*&---------------------------------------------------------------------*
FORM f_file_open.
  DATA: lt_file_table TYPE filetable,
        lv_rc         TYPE i,
        lv_user_action TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING
      window_title            = 'Select Excel File'
      default_extension       = 'xlsx'
      file_filter             = 'Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*'
    CHANGING
      file_table              = lt_file_table
      rc                      = lv_rc
      user_action             = lv_user_action
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5
  ).

  IF sy-subrc = 0 AND lv_user_action <> cl_gui_frontend_services=>action_cancel.
    READ TABLE lt_file_table INTO DATA(ls_file) INDEX 1.
    IF sy-subrc = 0.
      p_file = ls_file-filename.
    ENDIF.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_upload_excel
*&---------------------------------------------------------------------*
FORM f_upload_excel.
  DATA: lt_raw_data TYPE truxs_t_text_data,
        lv_filename TYPE rlgrap-filename.

  lv_filename = p_file.

  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_line_header        = 'X'
      i_tab_raw_data       = lt_raw_data
      i_filename           = lv_filename
    TABLES
      i_tab_converted_data = gt_excel_raw
    EXCEPTIONS
      conversion_failed    = 1
      OTHERS               = 2.

  IF sy-subrc <> 0.
    MESSAGE 'Error uploading file' TYPE 'E'.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_process_data
*&---------------------------------------------------------------------*
FORM f_process_data.
  DATA: ls_excel       LIKE LINE OF gt_excel_raw,
        lv_hex_string  TYPE string,
        lv_clean_hex   TYPE string,
        lv_xstring     TYPE xstring,
        lv_xml_string  TYPE string,
        lo_conv        TYPE REF TO cl_abap_conv_in_ce,
        lv_len         TYPE i,
        lv_idx         TYPE i,
        lv_char        TYPE c.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, lv_hex_string, lv_clean_hex, lv_xstring, lv_xml_string.
    
    gs_final-excel_id = ls_excel-col_a.
    gs_final-col_c    = ls_excel-col_c.

    " 1. Identify Hex Column (Assuming it's mostly in Col B based on previous context)
    "    Also checking Col C just in case.
    IF strlen( ls_excel-col_b ) > 10.
      lv_hex_string = ls_excel-col_b.
    ELSEIF strlen( ls_excel-col_c ) > 10.
      lv_hex_string = ls_excel-col_c.
    ENDIF.

    " 2. Clean Hex String manually (only 0-9, A-F)
    lv_len = strlen( lv_hex_string ).
    DO lv_len TIMES.
      lv_idx = sy-index - 1.
      lv_char = lv_hex_string+lv_idx(1).
      IF lv_char CA '0123456789ABCDEFabcdef'.
        CONCATENATE lv_clean_hex lv_char INTO lv_clean_hex.
      ENDIF.
    ENDDO.

    IF lv_clean_hex IS INITIAL.
      gs_final-error_msg = 'No Hex Data'.
      APPEND gs_final TO gt_final.
      CONTINUE.
    ENDIF.

    " 3. Convert Hex -> XString
    TRY.
        lv_xstring = lv_clean_hex.
      CATCH cx_root.
        gs_final-error_msg = 'Hex->XString Conversion Failed'.
        APPEND gs_final TO gt_final.
        CONTINUE.
    ENDTRY.

    " 4. Convert XString -> String (UTF-8)
    TRY.
        lo_conv = cl_abap_conv_in_ce=>create(
                    encoding    = 'UTF-8'
                    replacement = '?'
                    ignore_cerr = 'X' ).

        lo_conv->read(
          EXPORTING data = lv_xstring
          IMPORTING data = lv_xml_string ).

        gs_final-xml_content = lv_xml_string.

      CATCH cx_root.
        gs_final-error_msg = 'UTF-8 Conversion Failed'.
        APPEND gs_final TO gt_final.
        CONTINUE.
    ENDTRY.

    APPEND gs_final TO gt_final.

  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_display_alv
*&---------------------------------------------------------------------*
FORM f_display_alv.
  DATA: lo_alv TYPE REF TO cl_salv_table,
        lo_msg TYPE REF TO cx_salv_msg,
        lo_cols TYPE REF TO cl_salv_columns_table,
        lo_col  TYPE REF TO cl_salv_column.

  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table      = gt_final
      ).

      lo_cols = lo_alv->get_columns( ).
      lo_cols->set_optimize( 'X' ).
      
      " Adjust column labels
      TRY.
          lo_col = lo_cols->get_column( 'XML_CONTENT' ).
          lo_col->set_long_text( 'Converted XML Content' ).
          lo_col->set_medium_text( 'XML Content' ).
          lo_col->set_short_text( 'XML' ).
          " Set to reasonable output length
          lo_col->set_output_length( 100 ). 
      CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
