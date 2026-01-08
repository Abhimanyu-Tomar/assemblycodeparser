*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Parse Pipe-Delimited String, Display ALV
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

*----------------------------------------------------------------------*
* Types Declaration
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_excel_raw,
         col_a TYPE string,      " Field 1 (ID)
         col_b TYPE string,      " Field 2 (Possibly Empty in some views)
         col_c TYPE string,      " Field 3 (Complex String)
       END OF ty_excel_raw.

* Final Output Structure for ALV
TYPES: BEGIN OF ty_final,
         id          TYPE string,      " From Col A
         " Parsed fields from string
         segment_1   TYPE string,
         segment_2   TYPE string,
         segment_3   TYPE string,
         segment_4   TYPE string,
         segment_5   TYPE string,
         suffix      TYPE string,
         " Raw data for reference
         raw_string  TYPE string,
       END OF ty_final.

*----------------------------------------------------------------------*
* Data Declaration
*----------------------------------------------------------------------*
DATA: gt_excel_raw TYPE TABLE OF ty_excel_raw,
      gt_final     TYPE TABLE OF ty_final,
      gs_final     TYPE ty_final,
      gv_file      TYPE string.

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
  PERFORM f_process_data.
  PERFORM f_display_alv.

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
  DATA: ls_excel    LIKE LINE OF gt_excel_raw,
        lv_raw      TYPE string,
        lt_segments TYPE TABLE OF string,
        lv_count    TYPE i.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, lv_raw, lt_segments.

    gs_final-id = ls_excel-col_a.

    " Determine which column holds the complex string
    " Based on the user screenshot, it might be in Col C if Col B is empty
    IF ls_excel-col_b IS NOT INITIAL AND ls_excel-col_b CA '|;'.
      lv_raw = ls_excel-col_b.
    ELSEIF ls_excel-col_c IS NOT INITIAL.
      lv_raw = ls_excel-col_c.
    ELSE.
      " Fallback: use Col B if Col C is empty, even if Col B doesn't look complex
      lv_raw = ls_excel-col_b.
    ENDIF.

    gs_final-raw_string = lv_raw.

    " Logic to parse the string: Split by Pipe (|)
    IF lv_raw IS NOT INITIAL.
      SPLIT lv_raw AT '|' INTO TABLE lt_segments.
      
      " Map segments to fields
      DESCRIBE TABLE lt_segments LINES lv_count.
      
      LOOP AT lt_segments INTO DATA(lv_segment).
        CASE sy-tabix.
          WHEN 1. gs_final-segment_1 = lv_segment.
          WHEN 2. gs_final-segment_2 = lv_segment.
          WHEN 3. gs_final-segment_3 = lv_segment.
          WHEN 4. gs_final-segment_4 = lv_segment.
          WHEN 5. gs_final-segment_5 = lv_segment.
          " The last segment often contains the suffix (Company Name etc)
          " If there are many segments, the last one might be special
        ENDCASE.
      ENDLOOP.
      
      " Optional: If the string has NO pipes, put it all in Segment 1
      IF lv_count = 0.
         gs_final-segment_1 = lv_raw.
      ENDIF.
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

      " Rename columns for better readability
      TRY.
          lo_col = lo_cols->get_column( 'SEGMENT_1' ).
          lo_col->set_long_text( 'Data Segment 1' ).
          lo_col->set_medium_text( 'Segment 1' ).
          lo_col->set_short_text( 'Seg 1' ).

          lo_col = lo_cols->get_column( 'SEGMENT_2' ).
          lo_col->set_long_text( 'Data Segment 2' ).

          lo_col = lo_cols->get_column( 'RAW_STRING' ).
          lo_col->set_long_text( 'Original String' ).
        CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
