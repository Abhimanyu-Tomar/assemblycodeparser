*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel (Robust Type Check), Hex->String, Pretty Print
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

* Output Structure
TYPES: BEGIN OF ty_final,
         col_a     TYPE string,
         converted TYPE string,      " Formatted XML
         col_c     TYPE string,
         status    TYPE string,
       END OF ty_final.

DATA: gt_final TYPE TABLE OF ty_final,
      gs_final TYPE ty_final,
      gv_file  TYPE string.

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
    READ TABLE lt_file_table INTO DATA(ls_file) INDEX 1.
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
        lv_hex_raw     TYPE string.

  FIELD-SYMBOLS: <lt_data> TYPE STANDARD TABLE,
                 <ls_row>  TYPE any,
                 <lv_val>  TYPE any.

  " 1. Read File to XSTRING
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
    TABLES
      binary_tab   = lt_bin
    IMPORTING
      buffer       = lv_xstring.

  " 2. Parse Excel using CL_FDT_XL_SPREADSHEET
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

      READ TABLE lt_worksheets INTO DATA(lv_sheet) INDEX 1.

      lr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                  worksheet_name = lv_sheet ).

      ASSIGN lr_data->* TO <lt_data>.

      LOOP AT <lt_data> INTO <ls_row>.
        IF sy-tabix = 1. CONTINUE. ENDIF. " Header

        CLEAR: gs_final, lv_hex_raw.
        
        " Map Columns safely (Catching Move Errors)
        ASSIGN COMPONENT 1 OF STRUCTURE <ls_row> TO <lv_val>.
        IF sy-subrc = 0.
          TRY. gs_final-col_a = <lv_val>. CATCH cx_root. ENDTRY.
        ENDIF.

        ASSIGN COMPONENT 2 OF STRUCTURE <ls_row> TO <lv_val>.
        IF sy-subrc = 0.
          TRY. lv_hex_raw = <lv_val>. CATCH cx_root. ENDTRY.
        ENDIF.

        ASSIGN COMPONENT 3 OF STRUCTURE <ls_row> TO <lv_val>.
        IF sy-subrc = 0.
           TRY.
             IF strlen( lv_hex_raw ) < 20 AND strlen( <lv_val> ) > 20.
               lv_hex_raw = <lv_val>. 
             ELSE.
               gs_final-col_c = <lv_val>.
             ENDIF.
           CATCH cx_root.
           ENDTRY.
        ENDIF.

        " 3. Process Hex
        PERFORM f_convert_hex USING lv_hex_raw CHANGING gs_final.
        
        APPEND gs_final TO gt_final.
      ENDLOOP.

    CATCH cx_root.
      MESSAGE 'Excel Parse Error' TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_convert_hex
*&---------------------------------------------------------------------*
FORM f_convert_hex USING pv_hex TYPE string CHANGING ps_out TYPE ty_final.
  DATA: lv_clean   TYPE string,
        lv_xstr    TYPE xstring,
        lv_xml_raw TYPE string,
        lv_xml_pretty TYPE string,
        lo_ixml    TYPE REF TO if_ixml,
        lo_sf      TYPE REF TO if_ixml_stream_factory,
        lo_doc     TYPE REF TO if_ixml_document,
        lo_parser  TYPE REF TO if_ixml_parser,
        lo_render  TYPE REF TO if_ixml_renderer,
        lo_out     TYPE REF TO if_ixml_ostream,
        lo_in      TYPE REF TO if_ixml_istream,
        lo_err     TYPE REF TO if_ixml_parse_error.

  IF pv_hex IS INITIAL. RETURN. ENDIF.

  " Clean
  lv_clean = pv_hex.
  REPLACE ALL OCCURRENCES OF REGEX '[^0-9A-Fa-f]' IN lv_clean WITH ''.
  
  IF lv_clean IS INITIAL. ps_out-status = 'No Valid Hex'. RETURN. ENDIF.

  TRY.
      lv_xstr = lv_clean.

      lo_ixml = cl_ixml=>create( ).
      lo_sf   = lo_ixml->create_stream_factory( ).
      
      " Parse Direct from XString (Robust)
      lo_doc = lo_ixml->create_document( ).
      lo_in  = lo_sf->create_istream_xstring( lv_xstr ).
      lo_parser = lo_ixml->create_parser( stream_factory = lo_sf
                                          istream        = lo_in
                                          document       = lo_doc ).
      
      IF lo_parser->parse( ) = 0.
        lo_out = lo_sf->create_ostream_cstring( lv_xml_pretty ).
        lo_render = lo_ixml->create_renderer( ostream = lo_out document = lo_doc ).
        lo_render->set_normalizing( 'X' ).
        lo_render->render( ).
        ps_out-converted = lv_xml_pretty.
        ps_out-status = 'Success'.
      ELSE.
        " Fallback: FM
        CALL FUNCTION 'HR_RU_CONVERT_HEX_TO_STRING'
          EXPORTING xstring = lv_xstr
          IMPORTING cstring = lv_xml_raw.
        ps_out-converted = lv_xml_raw.
        
        lo_err = lo_parser->get_error( index = 0 ).
        IF lo_err IS BOUND.
           ps_out-status = lo_err->get_reason( ).
        ELSE.
           ps_out-status = 'Parse Error'.
        ENDIF.
      ENDIF.

    CATCH cx_root.
      ps_out-status = 'Conversion Exception'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_display_alv
*&---------------------------------------------------------------------*
FORM f_display_alv.
  DATA: lo_alv TYPE REF TO cl_salv_table.
  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_alv CHANGING t_table = gt_final ).
      lo_alv->get_columns( )->set_optimize( 'X' ).
      
      " Increase column width for readability
      TRY.
        DATA(lo_col) = lo_alv->get_columns( )->get_column( 'CONVERTED' ).
        lo_col->set_output_length( 100 ).
      CATCH cx_root.
      ENDTRY.

      lo_alv->display( ).
    CATCH cx_root.
  ENDTRY.
ENDFORM.
