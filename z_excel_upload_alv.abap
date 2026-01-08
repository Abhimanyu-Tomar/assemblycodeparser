*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Decode Base64, Display ALV
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

*----------------------------------------------------------------------*
* Type Definitions
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_excel_raw,
         col_a TYPE string,      " Field 1
         col_b TYPE string,      " Field 2 (Encoded String)
         col_c TYPE string,      " Field 3
       END OF ty_excel_raw.

* Output Structure
TYPES: BEGIN OF ty_final,
         col_a     TYPE string,
         col_b     TYPE string,      " Decoded String (XML)
         col_c     TYPE string,
         error_msg TYPE string,
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
        lv_input_str   TYPE string,
        lv_decoded_str TYPE string.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, lv_input_str, lv_decoded_str.
    
    gs_final-col_a = ls_excel-col_a.
    gs_final-col_c = ls_excel-col_c.

    " Determine which column has the encoded data
    IF strlen( ls_excel-col_b ) > 10.
      lv_input_str = ls_excel-col_b.
    ELSEIF strlen( ls_excel-col_c ) > 10.
      lv_input_str = ls_excel-col_c.
    ENDIF.

    IF lv_input_str IS NOT INITIAL.
      " Clean input (remove newlines/spaces) just in case
      " CONDENSE lv_input_str NO-GAPS. " Caution: Might break Base64 padding if present?
      " Base64 relies on correct length, but whitespace is usually ignored by decoders.
      " Let's remove typical whitespace:
      REPLACE ALL OCCURRENCES OF REGEX '\s' IN lv_input_str WITH ''.

      TRY.
          " Use cl_http_utility=>decode_base64 as requested
          lv_decoded_str = cl_http_utility=>decode_base64( encoded = lv_input_str ).
          
          gs_final-col_b = lv_decoded_str.

        CATCH cx_root.
          gs_final-error_msg = 'Base64 Decode Failed'.
      ENDTRY.
    ELSE.
      gs_final-error_msg = 'No input data to decode'.
    ENDIF.

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
          lo_col = lo_cols->get_column( 'COL_B' ).
          lo_col->set_long_text( 'Decoded Content' ).
          lo_col->set_medium_text( 'Decoded' ).
          lo_col->set_short_text( 'XML' ).
          lo_col->set_output_length( 100 ). 
      CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
