*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Hex->String (Specific Encoding), Pretty Print
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
         col_a     TYPE string,
         converted TYPE string,      " Formatted XML
         col_c     TYPE string,
         status    TYPE string,
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
        lv_xml_raw     TYPE string,
        lv_xml_pretty  TYPE string,
        lv_len         TYPE i,
        lv_idx         TYPE i,
        lv_char        TYPE c,
        lo_conv_out    TYPE REF TO cl_abap_conv_out_ce,
        lo_conv_in     TYPE REF TO cl_abap_conv_in_ce,
        lo_ixml        TYPE REF TO if_ixml,
        lo_stream_factory TYPE REF TO if_ixml_stream_factory,
        lo_document    TYPE REF TO if_ixml_document,
        lo_parser      TYPE REF TO if_ixml_parser,
        lo_renderer    TYPE REF TO if_ixml_renderer,
        lo_istream     TYPE REF TO if_ixml_istream,
        lo_ostream     TYPE REF TO if_ixml_ostream,
        lv_rc          TYPE i.

  " Initialize iXML Factory
  lo_ixml = cl_ixml=>create( ).
  lo_stream_factory = lo_ixml->create_stream_factory( ).

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: gs_final, lv_hex_string, lv_clean_hex, lv_xstring, lv_xml_raw, lv_xml_pretty.
    
    gs_final-col_a = ls_excel-col_a.
    gs_final-col_c = ls_excel-col_c.

    " Identify Hex Column
    IF strlen( ls_excel-col_b ) > 10.
      lv_hex_string = ls_excel-col_b.
    ELSEIF strlen( ls_excel-col_c ) > 10.
      lv_hex_string = ls_excel-col_c.
    ENDIF.

    IF lv_hex_string IS NOT INITIAL.
      " Clean Hex String
      lv_len = strlen( lv_hex_string ).
      DO lv_len TIMES.
        lv_idx = sy-index - 1.
        lv_char = lv_hex_string+lv_idx(1).
        IF lv_char CA '0123456789ABCDEFabcdef'.
          CONCATENATE lv_clean_hex lv_char INTO lv_clean_hex.
        ENDIF.
      ENDDO.

      IF lv_clean_hex IS NOT INITIAL.
        TRY.
            " 1. Hex -> XString (Buffer)
            lv_xstring = lv_clean_hex.

            " 2. Re-encoding logic as requested by user
            " Simulate user's snippet logic: String -> Buffer (1164) -> String (UTF-8)
            " But since we start with Hex (Buffer), we skip the first step or adapt it.
            " The user's snippet is fixing garbled TEXT. Our input IS Hex.
            " If the Hex represents the 'buffer' in the user's snippet, we just need the second part:
            " Buffer -> String (UTF-8).
            
            " However, if the user insists on the FULL snippet logic, maybe they think the Hex IS the 'text'?
            " That would be weird. Hex '3C3F' is not text 'Africa...'.
            " But let's assume the user wants exactly the conversion logic:
            
            " Scenario A: Hex IS the buffer (Correct interpretation of data flow)
            " We just need step 2 of their snippet:
            lo_conv_in = cl_abap_conv_in_ce=>create(
                           encoding    = 'UTF-8'
                           ignore_cerr = 'X' ).
                           
            lo_conv_in->convert(
              EXPORTING input = lv_xstring
              IMPORTING data  = lv_xml_raw ).

            IF lv_xml_raw IS NOT INITIAL.
              " 3. Pretty Print XML using iXML
              lo_document = lo_ixml->create_document( ).
              lo_istream  = lo_stream_factory->create_istream_string( lv_xml_raw ).
              lo_parser   = lo_ixml->create_parser( stream_factory = lo_stream_factory
                                                    istream        = lo_istream
                                                    document       = lo_document ).
              lv_rc = lo_parser->parse( ).

              IF lv_rc = 0.
                lo_ostream  = lo_stream_factory->create_ostream_cstring( lv_xml_pretty ).
                lo_renderer = lo_ixml->create_renderer( ostream  = lo_ostream
                                                        document = lo_document ).
                lo_renderer->set_normalizing( 'X' ).
                lo_renderer->render( ).
                
                gs_final-converted = lv_xml_pretty.
                gs_final-status    = 'Parsed & Formatted'.
              ELSE.
                gs_final-converted = lv_xml_raw.
                gs_final-status    = 'Raw XML (Parsing Failed)'.
              ENDIF.
            ELSE.
               gs_final-status = 'Empty Result'.
            ENDIF.

          CATCH cx_root.
            gs_final-status = 'Conversion Failed'.
        ENDTRY.
      ELSE.
        gs_final-status = 'No Valid Hex'.
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
      
      TRY.
          lo_col = lo_cols->get_column( 'CONVERTED' ).
          lo_col->set_long_text( 'Converted Content' ).
          lo_col->set_medium_text( 'Content' ).
          lo_col->set_output_length( 100 ). 
      CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
