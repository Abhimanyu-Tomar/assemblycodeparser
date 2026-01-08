*&---------------------------------------------------------------------*
*& Report Z_EXCEL_UPLOAD_ALV
*&---------------------------------------------------------------------*
*& Description: Upload Excel, Decode Hex XML, Parse Data, Display ALV
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
         raw_str     TYPE string,      " Parsed STR for reference
         error_msg   TYPE string,      " Debugging info
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
    IF gt_final IS INITIAL.
      MESSAGE 'Excel uploaded but no XML data could be extracted. Check format.' TYPE 'S' DISPLAY LIKE 'E'.
    ELSE.
      PERFORM f_display_alv.
    ENDIF.
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
        lv_xml_string TYPE string,
        lo_ixml       TYPE REF TO if_ixml,
        lo_stream_factory TYPE REF TO if_ixml_stream_factory,
        lo_istream    TYPE REF TO if_ixml_istream,
        lo_document   TYPE REF TO if_ixml_document,
        lo_parser     TYPE REF TO if_ixml_parser,
        lo_items      TYPE REF TO if_ixml_node_collection,
        lo_iterator   TYPE REF TO if_ixml_node_iterator,
        lo_node       TYPE REF TO if_ixml_node,
        lo_child_node TYPE REF TO if_ixml_node,
        lo_children   TYPE REF TO if_ixml_node_list,
        lv_item_id    TYPE string,
        lv_str_content TYPE string,
        lt_parts      TYPE TABLE OF string,
        lv_part       TYPE string,
        lv_key        TYPE string,
        lv_value      TYPE string,
        lo_conv       TYPE REF TO cl_abap_conv_in_ce,
        lv_found_items TYPE abap_bool.

  lo_ixml = cl_ixml=>create( ).
  lo_stream_factory = lo_ixml->create_stream_factory( ).

  LOOP AT gt_excel_raw INTO ls_excel.
    CLEAR: lv_hex_string, gs_final, lv_found_items.
    gs_final-excel_id = ls_excel-col_a.
    
    " Clean and Determine Hex Column
    " Remove whitespaces that might break conversion
    CONDENSE ls_excel-col_b NO-GAPS.
    CONDENSE ls_excel-col_c NO-GAPS.

    IF ls_excel-col_b IS NOT INITIAL AND strlen( ls_excel-col_b ) > 10.
      lv_hex_string = ls_excel-col_b.
    ELSEIF ls_excel-col_c IS NOT INITIAL AND strlen( ls_excel-col_c ) > 10.
      lv_hex_string = ls_excel-col_c.
    ENDIF.

    IF lv_hex_string IS INITIAL.
      gs_final-error_msg = 'No Hex Data Found'.
      APPEND gs_final TO gt_final.
      CONTINUE.
    ENDIF.

    " 1. Convert Hex String to XString (Binary)
    TRY.
        lv_xstring = lv_hex_string.
      CATCH cx_root.
        gs_final-error_msg = 'Hex Conversion Failed'.
        APPEND gs_final TO gt_final.
        CONTINUE. 
    ENDTRY.

    " 2. Convert XString to XML String (UTF-8)
    TRY.
        lo_conv = cl_abap_conv_in_ce=>create( input = lv_xstring encoding = 'UTF-8' ).
        lo_conv->read( IMPORTING data = lv_xml_string ).
      CATCH cx_root.
        gs_final-error_msg = 'UTF-8 Conversion Failed'.
        APPEND gs_final TO gt_final.
        CONTINUE.
    ENDTRY.

    " 3. Parse XML
    lo_document = lo_ixml->create_document( ).
    lo_istream = lo_stream_factory->create_istream_string( lv_xml_string ).
    lo_parser = lo_ixml->create_parser( stream_factory = lo_stream_factory
                                        istream        = lo_istream
                                        document       = lo_document ).
    
    IF lo_parser->parse( ) <> 0.
       gs_final-error_msg = 'XML Parse Failed'.
       APPEND gs_final TO gt_final.
       CONTINUE.
    ENDIF.

    " 4. Extract <item> elements
    lo_items = lo_document->get_elements_by_tag_name( name = 'item' ).
    IF lo_items IS BOUND.
        lo_iterator = lo_items->create_iterator( ).
        lo_node = lo_iterator->get_next( ).

        WHILE lo_node IS BOUND.
          lv_found_items = abap_true.
          CLEAR: gs_final-xml_item_id, gs_final-raw_str, 
                 gs_final-ewoid, gs_final-status, gs_final-shorttext.
          
          gs_final-excel_id = ls_excel-col_a.

          " Get Children of <item> (ID and STR)
          lo_children = lo_node->get_children( ).
          DO lo_children->get_length( ) TIMES.
            lo_child_node = lo_children->get_item( index = sy-index - 1 ).
            CASE lo_child_node->get_name( ).
              WHEN 'ID'.
                lv_item_id = lo_child_node->get_value( ).
              WHEN 'STR'.
                lv_str_content = lo_child_node->get_value( ).
            ENDCASE.
          ENDDO.

          gs_final-xml_item_id = lv_item_id.
          gs_final-raw_str     = lv_str_content.

          " 5. Parse STR content: Key==Value||Key==Value
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
          lo_node = lo_iterator->get_next( ).
        ENDWHILE.
    ENDIF.
    
    IF lv_found_items = abap_false.
        gs_final-error_msg = 'No Items found in XML'.
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

      " Set Column Headers
      TRY.
        lo_col = lo_cols->get_column( 'RAW_STR' ).
        lo_col->set_visible( ' ' ). " Hide raw string if too long
      CATCH cx_salv_not_found.
      ENDTRY.

      lo_alv->display( ).

    CATCH cx_salv_msg INTO lo_msg.
      MESSAGE lo_msg TYPE 'E'.
  ENDTRY.
ENDFORM.
