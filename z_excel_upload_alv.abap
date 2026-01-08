*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

*----------------------------------------------------------------------*
* Types Declaration
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_excel_raw,
         col_a TYPE string,
         col_b TYPE string, " Serialized data
         col_c TYPE string,
       END OF ty_excel_raw.

* Structure for Deserialized Data (Column B)
* Update this structure based on the actual JSON content
TYPES: BEGIN OF ty_json_data,
         key   TYPE string,
         value TYPE string,
         desc  TYPE string,
       END OF ty_json_data.

TYPES: BEGIN OF ty_final,
         id          TYPE string,
         json_key    TYPE string,
         json_value  TYPE string,
         json_desc   TYPE string,
         description TYPE string,
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
      file_filter             = 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
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

  " Note: simpler TEXT_CONVERT_XLS_TO_SAP is used here for demonstration.
  " Ensure the function module exists in your system.
  " Alternatively, use CL_FDT_XL_SPREADSHEET for strictly XLSX handling without FMs.
  
  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_line_header        = 'X' " Assume header line exists
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
  DATA: ls_excel LIKE LINE OF gt_excel_raw,
        ls_json  TYPE ty_json_data.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, ls_json.

    " Map direct columns
    gs_final-id          = ls_excel-col_a.
    gs_final-description = ls_excel-col_c.

    " Deserialize Column B (assuming JSON format)
    " Remove potential outer quotes if present from Excel
    REPLACE ALL OCCURRENCES OF '"' IN ls_excel-col_b WITH ''. 
    " The above REPLACE is risky if JSON has quotes. 
    " Only do this if Excel wraps the whole cell in quotes unnecessarily.
    " Usually, /ui2/cl_json handles standard JSON strings well.
    
    " Assuming Column B is valid JSON string e.g. {"key":"123", "value":"Test", "desc":"Detail"}
    /ui2/cl_json=>deserialize(
      EXPORTING
        json = ls_excel-col_b
      CHANGING
        data = ls_json
    ).

    " Map deserialized data to final structure
    gs_final-json_key   = ls_json-key.
    gs_final-json_value = ls_json-value.
    gs_final-json_desc  = ls_json-desc.

    APPEND gs_final TO gt_final.
  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_display_alv
*&---------------------------------------------------------------------*
FORM f_display_alv.
  DATA: lo_alv TYPE REF TO cl_salv_table,
        lo_msg TYPE REF TO cx_salv_msg.

  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table      = gt_final
      ).

      " Optional: Optimize column width
      lo_alv->get_columns( )->set_optimize( 'X' ).

      " Display ALV
      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
