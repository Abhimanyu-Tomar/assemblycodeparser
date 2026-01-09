*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Convert HEX to XML, Display Complete XML
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

* Output Structure - Simple: Show complete XML
TYPES: BEGIN OF ty_output,
         record_id    TYPE string,      " Excel Record ID
         item_id      TYPE string,      " Item ID from XML
         complete_xml TYPE string,      " Complete XML content
         status       TYPE string,      " Conversion status
       END OF ty_output.

DATA: gt_output TYPE TABLE OF ty_output,
      gs_output TYPE ty_output,
      gv_file   TYPE string.

SELECTION-SCREEN BEGIN OF BLOCK b1.
  PARAMETERS: p_file TYPE localfile OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f_file_open.

START-OF-SELECTION.
  gv_file = p_file.
  PERFORM f_process_file.
  PERFORM f_display_alv.

*&---------------------------------------------------------------------*
*& Form f_file_open
*&---------------------------------------------------------------------*
FORM f_file_open.
  DATA: lt_file_table TYPE filetable,
        ls_file       TYPE file_table,
        lv_rc         TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING
      window_title      = 'Select Excel File'
      default_extension = 'xlsx'
      file_filter       = 'Excel Files (*.xlsx)|*.xlsx'
    CHANGING
      file_table        = lt_file_table
      rc                = lv_rc
    EXCEPTIONS OTHERS   = 1 ).

  IF sy-subrc = 0 AND lines( lt_file_table ) > 0.
    READ TABLE lt_file_table INTO ls_file INDEX 1.
    p_file = ls_file-filename.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_process_file
*&---------------------------------------------------------------------*
FORM f_process_file.
  DATA: lt_bin         TYPE solix_tab,
        lv_xstring     TYPE xstring,
        lv_len         TYPE i,
        lo_excel       TYPE REF TO cl_fdt_xl_spreadsheet,
        lt_worksheets  TYPE if_fdt_doc_spreadsheet=>t_worksheet_names,
        lr_data        TYPE REF TO data,
        lv_hex_raw     TYPE string,
        lv_sheet       TYPE string,
        lv_temp        TYPE string,
        lv_record_id   TYPE string,
        lo_error       TYPE REF TO cx_root,
        lv_error_text  TYPE string.

  FIELD-SYMBOLS: <lt_data> TYPE STANDARD TABLE,
                 <ls_row>  TYPE any,
                 <lv_val>  TYPE any.

  " Read File
  cl_gui_frontend_services=>gui_upload(
    EXPORTING
      filename   = gv_file
      filetype   = 'BIN'
    IMPORTING
      filelength = lv_len
    CHANGING
      data_tab   = lt_bin
    EXCEPTIONS OTHERS = 1 ).

  IF sy-subrc <> 0.
    MESSAGE 'Upload Failed' TYPE 'S' DISPLAY LIKE 'E'.
    RETURN.
  ENDIF.

  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_len
    IMPORTING
      buffer       = lv_xstring
    TABLES
      binary_tab   = lt_bin.

  " Parse Excel
  TRY.
      CREATE OBJECT lo_excel
        EXPORTING
          document_name = gv_file
          xdocument     = lv_xstring.

      lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
        IMPORTING worksheet_names = lt_worksheets ).

      IF lt_worksheets IS INITIAL.
        MESSAGE 'No worksheets found' TYPE 'S' DISPLAY LIKE 'E'.
        RETURN.
      ENDIF.

      READ TABLE lt_worksheets INTO lv_sheet INDEX 1.

      lr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                  worksheet_name = lv_sheet ).

      ASSIGN lr_data->* TO <lt_data>.

      LOOP AT <lt_data> ASSIGNING <ls_row>.
        IF sy-tabix = 1. CONTINUE. ENDIF.

        CLEAR: lv_hex_raw, lv_temp, lv_record_id.

        " Column 1
        ASSIGN COMPONENT 1 OF STRUCTURE <ls_row> TO <lv_val>.
        IF sy-subrc = 0 AND <lv_val> IS ASSIGNED.
          TRY.
              lv_record_id = <lv_val>.
              CONDENSE lv_record_id.
            CATCH cx_root.
          ENDTRY.
        ENDIF.

        " Column 2
        ASSIGN COMPONENT 2 OF STRUCTURE <ls_row> TO <lv_val>.
        IF sy-subrc = 0 AND <lv_val> IS ASSIGNED.
          TRY.
              lv_hex_raw = <lv_val>.
              CONDENSE lv_hex_raw NO-GAPS.
            CATCH cx_root.
          ENDTRY.
        ENDIF.

        " Column 3
        ASSIGN COMPONENT 3 OF STRUCTURE <ls_row> TO <lv_val>.
        IF sy-subrc = 0 AND <lv_val> IS ASSIGNED.
          TRY.
              lv_temp = <lv_val>.
              CONDENSE lv_temp NO-GAPS.

              IF strlen( lv_hex_raw ) < 50 AND strlen( lv_temp ) > 50.
                lv_hex_raw = lv_temp.
              ENDIF.
            CATCH cx_root.
          ENDTRY.
        ENDIF.

        PERFORM f_convert_and_parse
          USING lv_record_id lv_hex_raw.

      ENDLOOP.

    CATCH cx_root INTO lo_error.
      lv_error_text = lo_error->get_text( ).
      MESSAGE lv_error_text TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_convert_and_parse
*&---------------------------------------------------------------------*
FORM f_convert_and_parse
  USING pv_record_id TYPE string
        pv_hex       TYPE string.

  DATA: lv_clean      TYPE string,
        lv_xstr       TYPE xstring,
        lv_xml_string TYPE string,
        lv_len        TYPE i,
        lv_offset     TYPE i,
        lv_hex_pair   TYPE c LENGTH 2,
        lv_bin_tab    TYPE x LENGTH 1,
        lv_search_pos TYPE i,
        lv_end_pos    TYPE i,
        lv_id_value   TYPE string,
        lv_str_value  TYPE string,
        lv_item_block TYPE string,
        lv_remaining  TYPE string,
        lo_ex         TYPE REF TO cx_root,
        lo_conv       TYPE REF TO cl_abap_conv_in_ce.

  IF pv_hex IS INITIAL. RETURN. ENDIF.

  TRY.
      " Clean hex
      lv_clean = pv_hex.
      TRANSLATE lv_clean TO UPPER CASE.
      REPLACE ALL OCCURRENCES OF REGEX '[^0-9A-F]' IN lv_clean WITH ''.

      IF lv_clean IS INITIAL. RETURN. ENDIF.

      " Ensure even length
      lv_len = strlen( lv_clean ).
      IF lv_len MOD 2 <> 0.
        CONCATENATE '0' lv_clean INTO lv_clean.
        lv_len = lv_len + 1.
      ENDIF.

      " Convert to xstring
      CLEAR lv_xstr.
      
      " Optimized hex to xstring conversion
      TRY.
          lv_xstr = lv_clean.
        CATCH cx_sy_conversion_error.
          " Fallback to manual loop if direct assignment fails
          lv_offset = 0.
          WHILE lv_offset < lv_len.
            lv_hex_pair = lv_clean+lv_offset(2).
            lv_bin_tab = lv_hex_pair.
            CONCATENATE lv_xstr lv_bin_tab INTO lv_xstr IN BYTE MODE.
            lv_offset = lv_offset + 2.
          ENDWHILE.
      ENDTRY.

      " Convert to string using CL_ABAP_CONV_IN_CE
      TRY.
          lo_conv = cl_abap_conv_in_ce=>create( 
                      input = lv_xstr 
                      encoding = 'UTF-8' 
                      replacement = '?' 
                      ignore_cerr = abap_true ).
          
          lo_conv->read( IMPORTING data = lv_xml_string ).
        CATCH cx_root.
          RETURN.
      ENDTRY.

      IF lv_xml_string IS INITIAL. RETURN. ENDIF.

      " Parse items - Extract ID and complete STR content
      lv_remaining = lv_xml_string.

      DO 1000 TIMES.
        FIND '<item>' IN lv_remaining MATCH OFFSET lv_search_pos.
        IF sy-subrc <> 0. EXIT. ENDIF.

        FIND '</item>' IN lv_remaining MATCH OFFSET lv_end_pos.
        IF sy-subrc <> 0. EXIT. ENDIF.

        " Include length of closing tag
        lv_len = lv_end_pos - lv_search_pos + 7. 

        IF lv_len > 0 AND lv_len <= strlen( lv_remaining ).
          lv_item_block = lv_remaining+lv_search_pos(lv_len).
        ELSE.
          EXIT.
        ENDIF.

        CLEAR: lv_id_value, lv_str_value.

        " Extract ID
        PERFORM f_extract_tag
          USING lv_item_block 'ID'
          CHANGING lv_id_value.

        " Extract complete STR content
        PERFORM f_extract_tag_complete
          USING lv_item_block 'STR'
          CHANGING lv_str_value.

        " Add to output
        IF lv_id_value IS NOT INITIAL OR lv_str_value IS NOT INITIAL.
          CLEAR gs_output.
          gs_output-record_id = pv_record_id.
          gs_output-item_id = lv_id_value.
          gs_output-complete_xml = lv_str_value.
          gs_output-status = 'Success'.
          APPEND gs_output TO gt_output.
        ENDIF.

        IF lv_end_pos < strlen( lv_remaining ).
          lv_remaining = lv_remaining+lv_end_pos.
        ELSE.
          EXIT.
        ENDIF.
      ENDDO.

    CATCH cx_root INTO lo_ex.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_extract_tag
*&---------------------------------------------------------------------*
FORM f_extract_tag
  USING pv_xml     TYPE string
        pv_tag     TYPE string
  CHANGING pv_value TYPE string.

  DATA: lv_start_tag TYPE string,
        lv_end_tag   TYPE string,
        lv_start_pos TYPE i,
        lv_end_pos   TYPE i,
        lv_length    TYPE i.

  CONCATENATE '<' pv_tag '>' INTO lv_start_tag.
  CONCATENATE '</' pv_tag '>' INTO lv_end_tag.

  FIND lv_start_tag IN pv_xml MATCH OFFSET lv_start_pos.
  IF sy-subrc = 0.
    lv_start_pos = lv_start_pos + strlen( lv_start_tag ).

    FIND lv_end_tag IN pv_xml MATCH OFFSET lv_end_pos.
    IF sy-subrc = 0.
      lv_length = lv_end_pos - lv_start_pos.
      IF lv_length > 0.
        pv_value = pv_xml+lv_start_pos(lv_length).
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_extract_tag_complete
*&---------------------------------------------------------------------*
FORM f_extract_tag_complete
  USING pv_xml     TYPE string
        pv_tag     TYPE string
  CHANGING pv_value TYPE string.

  DATA: lv_start_tag TYPE string,
        lv_end_tag   TYPE string,
        lv_start_pos TYPE i,
        lv_end_pos   TYPE i,
        lv_temp      TYPE string,
        lv_xml_len   TYPE i,
        lv_temp_len  TYPE i.

  CLEAR pv_value.

  CONCATENATE '<' pv_tag '>' INTO lv_start_tag.
  CONCATENATE '</' pv_tag '>' INTO lv_end_tag.

  lv_xml_len = strlen( pv_xml ).

  FIND lv_start_tag IN pv_xml MATCH OFFSET lv_start_pos.
  IF sy-subrc = 0.
    lv_start_pos = lv_start_pos + strlen( lv_start_tag ).

    IF lv_start_pos < lv_xml_len.
      lv_temp = pv_xml+lv_start_pos.
      lv_temp_len = strlen( lv_temp ).

      FIND lv_end_tag IN lv_temp MATCH OFFSET lv_end_pos.
      IF sy-subrc = 0 AND lv_end_pos > 0.
        pv_value = lv_temp(lv_end_pos).
      ELSEIF sy-subrc <> 0.
        pv_value = lv_temp.
      ENDIF.
    ENDIF.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_display_alv
*&---------------------------------------------------------------------*
FORM f_display_alv.
  DATA: lo_alv       TYPE REF TO cl_salv_table,
        lo_columns   TYPE REF TO cl_salv_columns_table,
        lo_column    TYPE REF TO cl_salv_column_table,
        lo_functions TYPE REF TO cl_salv_functions_list,
        lo_display   TYPE REF TO cl_salv_display_settings.

  IF gt_output IS INITIAL.
    MESSAGE 'No data to display' TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table = gt_output ).

      lo_functions = lo_alv->get_functions( ).
      lo_functions->set_all( abap_true ).

      lo_display = lo_alv->get_display_settings( ).
      lo_display->set_striped_pattern( abap_true ).
      lo_display->set_list_header( 'Complete XML Data Extraction' ).

      lo_columns = lo_alv->get_columns( ).
      lo_columns->set_optimize( abap_true ).

      " Configure columns
      TRY.
          lo_column ?= lo_columns->get_column( 'RECORD_ID' ).
          lo_column->set_long_text( 'Record ID' ).
          lo_column->set_medium_text( 'Record ID' ).
          lo_column->set_short_text( 'Rec ID' ).
          lo_column->set_output_length( 12 ).
        CATCH cx_salv_not_found.
      ENDTRY.

      TRY.
          lo_column ?= lo_columns->get_column( 'ITEM_ID' ).
          lo_column->set_long_text( 'Item ID' ).
          lo_column->set_medium_text( 'Item ID' ).
          lo_column->set_short_text( 'Item' ).
          lo_column->set_output_length( 15 ).
        CATCH cx_salv_not_found.
      ENDTRY.

      TRY.
          lo_column ?= lo_columns->get_column( 'COMPLETE_XML' ).
          lo_column->set_long_text( 'Complete XML Content' ).
          lo_column->set_medium_text( 'XML Content' ).
          lo_column->set_short_text( 'XML' ).
          " lo_column->set_output_length( 200 ).
        CATCH cx_salv_not_found.
      ENDTRY.

      TRY.
          lo_column ?= lo_columns->get_column( 'STATUS' ).
          lo_column->set_long_text( 'Status' ).
          lo_column->set_medium_text( 'Status' ).
          lo_column->set_short_text( 'Status' ).
          lo_column->set_output_length( 15 ).
        CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_root INTO DATA(lo_error).
      MESSAGE lo_error->get_text( ) TYPE 'I'.
  ENDTRY.
ENDFORM.
