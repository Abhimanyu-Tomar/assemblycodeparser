*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Robust Parsing with Debug Info
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

*----------------------------------------------------------------------*
* Types Declaration
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_excel_raw,
         col_a TYPE string,      " Field 1
         col_b TYPE string,      " Field 2 (Hex String)
         col_c TYPE string,      " Field 3
       END OF ty_excel_raw.

* Final Output Structure for ALV
TYPES: BEGIN OF ty_final,
         excel_id    TYPE string,      " From Col A
         xml_item_id TYPE string,      " From XML <ID>
         ewoid       TYPE string,
         status      TYPE string,
         level       TYPE string,
         shorttext   TYPE string,
         decision_l2 TYPE string,
         score_r     TYPE string,
         score_g     TYPE string,
         hex_len     TYPE i,           " Debug: Length of hex string
         debug_info  TYPE string,      " Debug: First 50 chars of decoded string
         error_msg   TYPE string,      " Error details
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
  DATA: ls_excel      LIKE LINE OF gt_excel_raw,
        lv_hex_string TYPE string,
        lv_xstring    TYPE xstring,
        lo_conv       TYPE REF TO cl_abap_conv_in_ce,
        lv_full_xml_str TYPE string,
        lv_str_content TYPE string,
        lt_parts      TYPE TABLE OF string,
        lv_part       TYPE string,
        lv_key        TYPE string,
        lv_value      TYPE string,
        lv_found_items TYPE abap_bool,
        lv_match_off  TYPE i,
        lv_match_len  TYPE i,
        lv_sub_off    TYPE i,
        lv_sub_len    TYPE i.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: lv_hex_string, gs_final, lv_found_items, lv_xstring, lv_full_xml_str.
    gs_final-excel_id = ls_excel-col_a.
    
    " Clean and Determine Hex Column
    CONDENSE ls_excel-col_b NO-GAPS.
    CONDENSE ls_excel-col_c NO-GAPS.

    IF ls_excel-col_b IS NOT INITIAL AND strlen( ls_excel-col_b ) > 10.
      lv_hex_string = ls_excel-col_b.
    ELSEIF ls_excel-col_c IS NOT INITIAL AND strlen( ls_excel-col_c ) > 10.
      lv_hex_string = ls_excel-col_c.
    ENDIF.

    gs_final-hex_len = strlen( lv_hex_string ).

    IF lv_hex_string IS INITIAL.
      gs_final-error_msg = 'No Hex Data'.
      APPEND gs_final TO gt_final.
      CONTINUE.
    ENDIF.

    " 1. Convert Hex String to XString
    TRY.
        lv_xstring = lv_hex_string.
      CATCH cx_root.
        gs_final-error_msg = 'Hex Convert Fail'.
        APPEND gs_final TO gt_final.
        CONTINUE. 
    ENDTRY.

    " 2. Convert to String (Try UTF-8 first)
    TRY.
        lo_conv = cl_abap_conv_in_ce=>create( input = lv_xstring encoding = 'UTF-8' replacement = '?' ignore_cerr = 'X' ).
        lo_conv->read( IMPORTING data = lv_full_xml_str ).
      CATCH cx_root.
        gs_final-error_msg = 'UTF-8 Fail'.
        APPEND gs_final TO gt_final.
        CONTINUE.
    ENDTRY.

    " Populate Debug Info
    IF strlen( lv_full_xml_str ) > 50.
      gs_final-debug_info = substring( val = lv_full_xml_str len = 50 ).
    ELSE.
      gs_final-debug_info = lv_full_xml_str.
    ENDIF.

    " 3. Robust Search for <STR>...</STR>
    " We loop searching for <STR> to handle multiple occurrences
    lv_sub_off = 0.
    WHILE 1 = 1.
      FIND REGEX '<STR>(.*?)</STR>' IN SECTION OFFSET lv_sub_off OF lv_full_xml_str 
           MATCH OFFSET lv_match_off 
           MATCH LENGTH lv_match_len
           IGNORING CASE.
      
      IF sy-subrc <> 0.
        EXIT. 
      ENDIF.

      lv_found_items = abap_true.
      
      " Extract Content (remove tags)
      " Length of <STR> is 5, </STR> is 6. Total 11 chars overhead.
      " Content starts at MatchOffset + 5
      lv_str_content = substring( val = lv_full_xml_str off = lv_match_off + 5 len = lv_match_len - 11 ).
      
      " Reset fields
      CLEAR: gs_final-ewoid, gs_final-status, gs_final-shorttext,
             gs_final-decision_l2, gs_final-score_r, gs_final-score_g.
      
      " Attempt to find ID just before this STR (Optional, but good for completeness)
      " Simple lookback or just skipping it for now to ensure STR works first.

      " 4. Parse STR content
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
          WHEN 'SCORE_YES_R'. IF gs_final-score_r IS INITIAL. gs_final-score_r = lv_value. ENDIF.
          WHEN 'SCORE_YES_G'. IF gs_final-score_g IS INITIAL. gs_final-score_g = lv_value. ENDIF.
        ENDCASE.
      ENDLOOP.

      APPEND gs_final TO gt_final.
      
      " Advance offset
      lv_sub_off = lv_match_off + lv_match_len.
    ENDWHILE.
    
    IF lv_found_items = abap_false.
        gs_final-error_msg = 'No <STR> tags found'.
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

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
