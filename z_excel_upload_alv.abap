*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Convert Hex using HR_RU_CONVERT_HEX_TO_STRING
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
         xml_item_id TYPE string,
         ewoid       TYPE string,
         status      TYPE string,
         level       TYPE string,
         shorttext   TYPE string,
         decision_l2 TYPE string,
         score_r     TYPE string,
         score_g     TYPE string,
         raw_str     TYPE string,      " Decoded XML String (content)
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
        lv_xstring     TYPE xstring,
        lv_xml_string  TYPE string,
        lv_sub_off     TYPE i,
        lv_match_off   TYPE i,
        lv_match_len   TYPE i,
        lv_found_items TYPE abap_bool,
        lv_str_content TYPE string,
        lt_parts       TYPE TABLE OF string,
        lv_part        TYPE string,
        lv_key         TYPE string,
        lv_value       TYPE string.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, lv_hex_string, lv_xstring, lv_xml_string, lv_found_items.
    gs_final-excel_id = ls_excel-col_a.

    " Determine Hex Column & Clean
    " Based on User Logic: Keep only Valid Hex Chars [0-9A-Fa-f]
    IF ls_excel-col_b IS NOT INITIAL AND strlen( ls_excel-col_b ) > 10.
      lv_hex_string = ls_excel-col_b.
    ELSEIF ls_excel-col_c IS NOT INITIAL AND strlen( ls_excel-col_c ) > 10.
      lv_hex_string = ls_excel-col_c.
    ENDIF.

    REPLACE ALL OCCURRENCES OF REGEX '[^0-9A-Fa-f]' IN lv_hex_string WITH ''.

    IF lv_hex_string IS INITIAL.
      gs_final-error_msg = 'No Valid Hex Data'.
      APPEND gs_final TO gt_final.
      CONTINUE.
    ENDIF.

    " Convert Hex using HR Function Module as requested
    TRY.
        lv_xstring = lv_hex_string.

        CALL FUNCTION 'HR_RU_CONVERT_HEX_TO_STRING'
          EXPORTING
            xstring = lv_xstring
          IMPORTING
            cstring = lv_xml_string.

      CATCH cx_root.
        gs_final-error_msg = 'Hex Conversion Failed'.
        APPEND gs_final TO gt_final.
        CONTINUE.
    ENDTRY.

    " Parse the Decoded String (XML-like)
    " Robust Search for <STR>...</STR> blocks
    lv_sub_off = 0.
    WHILE 1 = 1.
      FIND REGEX '<STR>(.*?)</STR>' IN SECTION OFFSET lv_sub_off OF lv_xml_string
           MATCH OFFSET lv_match_off
           MATCH LENGTH lv_match_len
           IGNORING CASE.
      
      IF sy-subrc <> 0.
        EXIT. 
      ENDIF.

      lv_found_items = abap_true.
      
      " Extract Content: MatchOffset + 5 (len of <STR>), Length = MatchLen - 11 (len of <STR></STR>)
      lv_str_content = substring( val = lv_xml_string off = lv_match_off + 5 len = lv_match_len - 11 ).
      gs_final-raw_str = lv_str_content. " Save for reference

      " Reset fields for new item
      CLEAR: gs_final-ewoid, gs_final-status, gs_final-shorttext, gs_final-xml_item_id,
             gs_final-decision_l2, gs_final-score_r, gs_final-score_g.

      " Attempt to find ID for this block (Look backwards from STR pos)
      " Simple heuristic: Find last <ID> before this <STR>
      DATA: lv_id_match_off TYPE i, lv_id_match_len TYPE i.
      FIND REGEX '<ID>(.*?)</ID>' IN SECTION OFFSET 0 LENGTH lv_match_off OF lv_xml_string
           MATCH OFFSET lv_id_match_off
           MATCH LENGTH lv_id_match_len
           IGNORING CASE.
      IF sy-subrc = 0.
         " If multiple IDs exist, this finds the *first* one in the block. 
         " For strict XML matching, we'd need loop logic, but this often suffices for sequential data.
         " Better approach: Search in the small chunk between previous STR end and current STR start.
         gs_final-xml_item_id = substring( val = lv_xml_string off = lv_id_match_off + 4 len = lv_id_match_len - 9 ).
      ENDIF.

      " Parse Key==Value||...
      REPLACE ALL OCCURRENCES OF '&lt;' IN lv_str_content WITH '<'.
      REPLACE ALL OCCURRENCES OF '&gt;' IN lv_str_content WITH '>'.
      REPLACE ALL OCCURRENCES OF '&amp;' IN lv_str_content WITH '&'.

      SPLIT lv_str_content AT '||' INTO TABLE lt_parts.
      LOOP AT lt_parts INTO lv_part.
        SPLIT lv_part AT '==' INTO lv_key lv_value.
        CONDENSE lv_key.
        CASE lv_key.
          WHEN 'EWOID'.       gs_final-ewoid       = lv_value.
          WHEN 'STATUS'.      gs_final-status      = lv_value.
          WHEN 'LEVEL'.       gs_final-level       = lv_value.
          WHEN 'SHORTTEXT'.   gs_final-shorttext   = lv_value.
          WHEN 'DECISION_L2'. gs_final-decision_l2 = lv_value.
          WHEN 'L2_SCORE_R'.  gs_final-score_r     = lv_value.
          WHEN 'L2_SCORE_G'.  gs_final-score_g     = lv_value.
        ENDCASE.
      ENDLOOP.

      APPEND gs_final TO gt_final.
      lv_sub_off = lv_match_off + lv_match_len.
    ENDWHILE.

    IF lv_found_items = abap_false.
        gs_final-error_msg = 'No <STR> tags found in decoded content'.
        gs_final-raw_str = lv_xml_string(100). " Show first 100 chars
        APPEND gs_final TO gt_final.
    ENDIF.

  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_display_alv
*&---------------------------------------------------------------------*
FORM f_display_alv.
  DATA: lo_alv TYPE REF TO cl_salv_table,
        lo_msg TYPE REF TO cx_salv_msg,
        lo_cols TYPE REF TO cl_salv_columns_table.

  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table      = gt_final
      ).

      lo_cols = lo_alv->get_columns( ).
      lo_cols->set_optimize( 'X' ).

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
