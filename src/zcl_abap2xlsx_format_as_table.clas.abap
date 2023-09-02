CLASS zcl_abap2xlsx_format_as_table DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC.

  PUBLIC SECTION.
    METHODS constructor.
    METHODS main
      IMPORTING !screen_title TYPE csequence
      RAISING   zcx_excel.

  PROTECTED SECTION.
  PRIVATE SECTION.
    DATA file_handler TYPE REF TO lcl_file_handler.
    DATA ooxml_reader TYPE REF TO zif_excel_reader.
    DATA ooxml_writer TYPE REF TO zif_excel_writer.
    DATA workbook TYPE REF TO zcl_excel.
ENDCLASS.


CLASS zcl_abap2xlsx_format_as_table IMPLEMENTATION.

  METHOD constructor.
    file_handler = NEW lcl_file_handler( ).
  ENDMETHOD.

  METHOD main.
    DATA(selected_file) = file_handler->select_excel_file( )->get_file_name( ).

    FIND FIRST OCCURRENCE OF PCRE '(\.xlsx|\.xlsm)\s*$' IN selected_file SUBMATCHES DATA(extension).
    CASE to_upper( extension ).
      WHEN '.XLSX'.
        ooxml_reader = NEW zcl_excel_reader_2007( ).
        ooxml_writer = NEW zcl_excel_writer_2007( ).
      WHEN '.XLSM'.
        ooxml_reader = NEW zcl_excel_reader_xlsm( ).
        ooxml_writer = NEW zcl_excel_writer_xlsm( ).
      WHEN OTHERS.
        MESSAGE |Unsupported file extension '{ extension }'| TYPE 'E'.
    ENDCASE.

    workbook = ooxml_reader->load_file( selected_file ).
    file_handler->save_excel_file( workbook_content = ooxml_writer->write_file( io_excel = workbook )
                                   file_name        = |abap2xlsx_format_as_table{ extension }| ).
  ENDMETHOD.

ENDCLASS.
