*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Hex->XString->String (SCMS), Parse Data
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
         debug_hex   TYPE string,      " Debug: First 20 chars of Hex
         debug_xml   TYPE string,      " Debug: First 50 chars of XML
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
        lv_end_off     TYPE i,
        lv_str_start   TYPE i,
        lv_str_len     TYPE i,
        lv_found_items TYPE abap_bool,
        lv_str_content TYPE string,
        lt_parts       TYPE TABLE OF string,
        lv_part        TYPE string,
        lv_key         TYPE string,
        lv_value       TYPE string,
        lv_temp_off    TYPE i,
        lv_id_start    TYPE i,
        lv_id_end      TYPE i,
        lv_id_len      TYPE i,
        lv_last_id_start TYPE i,
        lv_last_id_end   TYPE i,
        lv_mimetype    TYPE string.

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, lv_hex_string, lv_xstring, lv_xml_string, lv_found_items.
    gs_final-excel_id = ls_excel-col_a.

    " 1. Identify and Clean Hex Column
    " Heuristic: Longest column likely contains the XML Hex
    IF strlen( ls_excel-col_b ) > strlen( ls_excel-col_c ) AND strlen( ls_excel-col_b ) > 10.
      lv_hex_string = ls_excel-col_b.
    ELSEIF strlen( ls_excel-col_c ) > 10.
      lv_hex_string = ls_excel-col_c.
    ENDIF.
    
    gs_final-debug_hex = substring( val = lv_hex_string len = 20 ).

    " Remove all non-hex characters (newlines, spaces, etc)
    REPLACE ALL OCCURRENCES OF REGEX '[^0-9A-Fa-f]' IN lv_hex_string WITH ''.

    IF lv_hex_string IS INITIAL.
      gs_final-error_msg = 'No Hex Data'.
      APPEND gs_final TO gt_final.
      CONTINUE.
    ENDIF.

    " 2. Convert Hex String to XString
    TRY.
        lv_xstring = lv_hex_string.
      CATCH cx_root.
        gs_final-error_msg = 'Hex Conversion Failed'.
        APPEND gs_final TO gt_final.
        CONTINUE. 
    ENDTRY.

    " 3. Convert XString to String using SCMS Function
    " This is highly robust and handles various encodings if needed
    CALL FUNCTION 'SCMS_XSTRING_TO_STRING'
      EXPORTING
        buffer        = lv_xstring
        encoding      = '4110' " UTF-8
      IMPORTING
        output_string = lv_xml_string
      EXCEPTIONS
        failed        = 1
        OTHERS        = 2.

    IF sy-subrc <> 0.
      " Fallback: Try with '1100' (ISO-8859-1) or no encoding
      CALL FUNCTION 'SCMS_XSTRING_TO_STRING'
        EXPORTING
          buffer        = lv_xstring
        IMPORTING
          output_string = lv_xml_string
        EXCEPTIONS
          OTHERS        = 1.
    ENDIF.

    " Populate debug info
    IF strlen( lv_xml_string ) > 50.
      gs_final-debug_xml = substring( val = lv_xml_string len = 50 ).
    ELSE.
      gs_final-debug_xml = lv_xml_string.
    ENDIF.

    " 4. Parse XML - Search for <STR> tags
    lv_sub_off = 0.
    
    WHILE 1 = 1.
      " Search for <STR> case-insensitive
      FIND FIRST OCCURRENCE OF '<STR>' IN SECTION OFFSET lv_sub_off OF lv_xml_string 
           MATCH OFFSET lv_match_off 
           IGNORING CASE.
      
      IF sy-subrc <> 0.
        EXIT. 
      ENDIF.
      
      lv_str_start = lv_match_off + 5. 

      " Search for closing </STR> case-insensitive
      FIND FIRST OCCURRENCE OF '</STR>' IN SECTION OFFSET lv_str_start OF lv_xml_string 
           MATCH OFFSET lv_end_off 
           IGNORING CASE.
      
      IF sy-subrc <> 0.
        EXIT. 
      ENDIF.

      lv_str_len = lv_end_off - lv_str_start.
      lv_found_items = abap_true.
      
      " Extract Content
      lv_str_content = substring( val = lv_xml_string off = lv_str_start len = lv_str_len ).

      " Reset fields
      CLEAR: gs_final-ewoid, gs_final-status, gs_final-shorttext, gs_final-xml_item_id,
             gs_final-decision_l2, gs_final-score_r, gs_final-score_g.
      
      " Find preceding ID
      CLEAR: lv_last_id_start, lv_last_id_end.
      lv_temp_off = 0.
      
      WHILE 1 = 1.
        DATA: lv_scope_len TYPE i.
        lv_scope_len = lv_match_off - lv_temp_off.
        
        IF lv_scope_len <= 0. EXIT. ENDIF.

        FIND FIRST OCCURRENCE OF '<ID>' IN SECTION OFFSET lv_temp_off LENGTH lv_scope_len OF lv_xml_string 
             MATCH OFFSET lv_id_start 
             IGNORING CASE.
             
        IF sy-subrc <> 0. EXIT. ENDIF.
        
        FIND FIRST OCCURRENCE OF '</ID>' IN SECTION OFFSET ( lv_id_start + 4 ) OF lv_xml_string 
             MATCH OFFSET lv_id_end 
             IGNORING CASE.
             
        IF sy-subrc <> 0 OR lv_id_end > lv_match_off. EXIT. ENDIF.

        lv_last_id_start = lv_id_start + 4.
        lv_last_id_end   = lv_id_end.
        lv_temp_off      = lv_id_end + 5.
      ENDWHILE.

      IF lv_last_id_start > 0.
         lv_id_len = lv_last_id_end - lv_last_id_start.
         gs_final-xml_item_id = substring( val = lv_xml_string off = lv_last_id_start len = lv_id_len ).
      ENDIF.

      " Parse Key==Value||...
      " Decode HTML entities
      REPLACE ALL OCCURRENCES OF '&lt;'   IN lv_str_content WITH '<' IGNORING CASE.
      REPLACE ALL OCCURRENCES OF '&gt;'   IN lv_str_content WITH '>' IGNORING CASE.
      REPLACE ALL OCCURRENCES OF '&amp;'  IN lv_str_content WITH '&' IGNORING CASE.
      REPLACE ALL OCCURRENCES OF '&quot;' IN lv_str_content WITH '"' IGNORING CASE.

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
      
      lv_sub_off = lv_end_off + 6. 
    ENDWHILE.

    IF lv_found_items = abap_false.
        gs_final-error_msg = 'No valid <STR> items found'.
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

      " Rename columns for better readability
      TRY.
          lo_col = lo_cols->get_column( 'DEBUG_HEX' ).
          lo_col->set_long_text( 'Hex Preview' ).
          lo_col = lo_cols->get_column( 'DEBUG_XML' ).
          lo_col->set_long_text( 'XML Preview' ).
      CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
