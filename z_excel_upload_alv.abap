*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Convert HEX to XML, Split || Data, Dynamic ALV
*&---------------------------------------------------------------------*
REPORT z_excel_upload_alv.

* Intermediate storage for parsed XML items
TYPES: BEGIN OF ty_raw_data,
         record_id   TYPE string,
         item_id     TYPE string,
         raw_content TYPE string,
       END OF ty_raw_data.

DATA: gt_raw_data TYPE TABLE OF ty_raw_data,
      gs_raw_data TYPE ty_raw_data,
      gv_file     TYPE string.

* Dynamic ALV Data
DATA: gr_alv_data TYPE REF TO data.

SELECTION-SCREEN BEGIN OF BLOCK b1.
  PARAMETERS: p_file TYPE localfile OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f_file_open.

START-OF-SELECTION.
  gv_file = p_file.
  PERFORM f_process_file.
  PERFORM f_prepare_and_display_alv.

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

        PERFORM f_convert_and_collect
          USING lv_record_id lv_hex_raw.

      ENDLOOP.

    CATCH cx_root INTO lo_error.
      lv_error_text = lo_error->get_text( ).
      MESSAGE lv_error_text TYPE 'S' DISPLAY LIKE 'E'.
  ENDTRY.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_convert_and_collect
*&---------------------------------------------------------------------*
FORM f_convert_and_collect
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
      TRY.
          lv_xstr = lv_clean.
        CATCH cx_sy_conversion_error.
          lv_offset = 0.
          WHILE lv_offset < lv_len.
            lv_hex_pair = lv_clean+lv_offset(2).
            lv_bin_tab = lv_hex_pair.
            CONCATENATE lv_xstr lv_bin_tab INTO lv_xstr IN BYTE MODE.
            lv_offset = lv_offset + 2.
          ENDWHILE.
      ENDTRY.

      " Convert to string
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

      lv_remaining = lv_xml_string.

      DO 1000 TIMES.
        IF lv_remaining IS INITIAL. EXIT. ENDIF.

        FIND '<item>' IN lv_remaining MATCH OFFSET lv_search_pos.
        IF sy-subrc <> 0. EXIT. ENDIF.

        FIND '</item>' IN lv_remaining MATCH OFFSET lv_end_pos.
        IF sy-subrc <> 0. EXIT. ENDIF.

        " Include length of closing tag (7 chars)
        lv_len = lv_end_pos - lv_search_pos + 7. 

        IF lv_len > 0 AND lv_len <= strlen( lv_remaining ).
          lv_item_block = lv_remaining+lv_search_pos(lv_len).
        ELSE.
          EXIT.
        ENDIF.

        CLEAR: lv_id_value, lv_str_value.

        PERFORM f_extract_tag
          USING lv_item_block 'ID'
          CHANGING lv_id_value.

        PERFORM f_extract_tag_complete
          USING lv_item_block 'STR'
          CHANGING lv_str_value.

        IF lv_id_value IS NOT INITIAL OR lv_str_value IS NOT INITIAL.
          CLEAR gs_raw_data.
          gs_raw_data-record_id = pv_record_id.
          gs_raw_data-item_id = lv_id_value.
          gs_raw_data-raw_content = lv_str_value.
          APPEND gs_raw_data TO gt_raw_data.
        ENDIF.

        " Advance remaining string past the current item
        lv_end_pos = lv_end_pos + 7.
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
  
  PERFORM f_extract_tag USING pv_xml pv_tag CHANGING pv_value.
ENDFORM.

*&---------------------------------------------------------------------*
*& Form f_prepare_and_display_alv
*&---------------------------------------------------------------------*
FORM f_prepare_and_display_alv.
  DATA: lt_fcat       TYPE lvc_t_fcat,
        ls_fcat       TYPE lvc_s_fcat,
        lo_alv        TYPE REF TO cl_salv_table,
        lt_splits     TYPE TABLE OF string,
        lv_max_cols   TYPE i,
        lv_col_name   TYPE string,
        lv_col_idx    TYPE i,
        ls_raw        LIKE LINE OF gt_raw_data,
        lv_split_val  TYPE string.

  FIELD-SYMBOLS: <lt_dyn_table> TYPE STANDARD TABLE,
                 <ls_dyn_line>  TYPE any,
                 <lv_field>     TYPE any.

  IF gt_raw_data IS INITIAL.
    MESSAGE 'No data extracted.' TYPE 'S' DISPLAY LIKE 'W'.
    RETURN.
  ENDIF.

  " 1. Determine maximum number of columns
  lv_max_cols = 0.
  LOOP AT gt_raw_data INTO ls_raw.
    SPLIT ls_raw-raw_content AT '||' INTO TABLE lt_splits.
    IF lines( lt_splits ) > lv_max_cols.
      lv_max_cols = lines( lt_splits ).
    ENDIF.
  ENDLOOP.

  IF lv_max_cols = 0. lv_max_cols = 1. ENDIF.

  " 2. Build Field Catalog
  " Fixed columns
  ls_fcat-fieldname = 'RECORD_ID'.
  ls_fcat-ref_table = 'NA'.
  ls_fcat-scrtext_s = 'RecID'.
  ls_fcat-scrtext_m = 'Record ID'.
  ls_fcat-scrtext_l = 'Record ID'.
  ls_fcat-inttype   = 'g'. " String
  ls_fcat-col_opt   = 'X'.
  APPEND ls_fcat TO lt_fcat.

  ls_fcat-fieldname = 'ITEM_ID'.
  ls_fcat-scrtext_s = 'ItemID'.
  ls_fcat-scrtext_m = 'Item ID'.
  ls_fcat-scrtext_l = 'Item ID'.
  ls_fcat-inttype   = 'g'.
  ls_fcat-col_opt   = 'X'.
  APPEND ls_fcat TO lt_fcat.

  " Dynamic columns
  DO lv_max_cols TIMES.
    lv_col_idx = sy-index.
    lv_col_name = |COL_{ lv_col_idx }|. 
    
    CLEAR ls_fcat.
    ls_fcat-fieldname = lv_col_name.
    ls_fcat-scrtext_s = |Col { lv_col_idx }|.
    ls_fcat-scrtext_m = |Column { lv_col_idx }|.
    ls_fcat-scrtext_l = |Data Column { lv_col_idx }|.
    ls_fcat-inttype   = 'g'. " String
    ls_fcat-col_opt   = 'X'.
    APPEND ls_fcat TO lt_fcat.
  ENDDO.

  " 3. Create Dynamic Table
  CALL METHOD cl_alv_table_create=>create_dynamic_table
    EXPORTING
      it_fieldcatalog = lt_fcat
    IMPORTING
      ep_table        = gr_alv_data.

  ASSIGN gr_alv_data->* TO <lt_dyn_table>.

  " 4. Fill Data
  LOOP AT gt_raw_data INTO ls_raw.
    APPEND INITIAL LINE TO <lt_dyn_table> ASSIGNING <ls_dyn_line>.

    " Fill Fixed Fields
    ASSIGN COMPONENT 'RECORD_ID' OF STRUCTURE <ls_dyn_line> TO <lv_field>.
    IF sy-subrc = 0. <lv_field> = ls_raw-record_id. ENDIF.

    ASSIGN COMPONENT 'ITEM_ID' OF STRUCTURE <ls_dyn_line> TO <lv_field>.
    IF sy-subrc = 0. <lv_field> = ls_raw-item_id. ENDIF.

    " Fill Dynamic Fields
    SPLIT ls_raw-raw_content AT '||' INTO TABLE lt_splits.
    
    LOOP AT lt_splits INTO lv_split_val.
      lv_col_idx = sy-tabix.
      lv_col_name = |COL_{ lv_col_idx }|.
      
      ASSIGN COMPONENT lv_col_name OF STRUCTURE <ls_dyn_line> TO <lv_field>.
      IF sy-subrc = 0.
        <lv_field> = lv_split_val.
      ENDIF.
    ENDLOOP.
  ENDLOOP.

  " 5. Display ALV
  TRY.
      cl_salv_table=>factory(
        IMPORTING
          r_salv_table = lo_alv
        CHANGING
          t_table = <lt_dyn_table> ).

      lo_alv->get_functions( )->set_all( abap_true ).
      lo_alv->get_columns( )->set_optimize( abap_true ).
      lo_alv->get_display_settings( )->set_striped_pattern( abap_true ).
      lo_alv->get_display_settings( )->set_list_header( 'Splitted XML Data' ).

      lo_alv->display( ).

    CATCH cx_root INTO DATA(lo_error).
      MESSAGE lo_error->get_text( ) TYPE 'I'.
  ENDTRY.
ENDFORM.
